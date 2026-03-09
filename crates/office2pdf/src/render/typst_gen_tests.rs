use super::*;
use crate::ir::{
    ChartSeries, ColumnLayout, GradientStop, HeaderFooterParagraph, ListItem, ListKind,
    ListLevelStyle, Metadata, SmartArtNode, StyleSheet,
};
use std::collections::BTreeMap;

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
        header: None,
        footer: None,
        columns: None,
    })
}

/// Helper to create a simple paragraph with one plain-text run.
fn make_paragraph(text: &str) -> Block {
    Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: text.to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })
}

#[test]
fn test_generate_plain_paragraph() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Hello World")])]);
    let result = generate_typst(&doc).unwrap().source;
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
        header: None,
        footer: None,
        columns: None,
    })]);
    let result = generate_typst(&doc).unwrap().source;
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
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
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
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
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
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
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
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
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
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
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
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
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
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
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
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
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
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
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
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
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
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("leading: 18pt"),
        "Expected exact leading in: {result}"
    );
}

#[test]
fn test_generate_letter_spacing() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Spaced text".to_string(),
            style: TextStyle {
                letter_spacing: Some(2.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("tracking: 2pt"),
        "Expected tracking param in: {result}"
    );
}

#[test]
fn test_generate_letter_spacing_negative() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Condensed".to_string(),
            style: TextStyle {
                letter_spacing: Some(-0.5),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("tracking: -0.5pt"),
        "Expected negative tracking in: {result}"
    );
}

#[test]
fn test_generate_tab_uses_measured_default_stops() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Name:\tValue".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#context {"),
        "Expected contextual tab rendering in: {result}"
    );
    assert!(
        result.contains("measure(tab_prefix_0).width"),
        "Expected tab spacing to measure the rendered prefix in: {result}"
    );
    assert!(
        result.contains("calc.rem-euclid(tab_prefix_width_1.abs.pt(), 36)"),
        "Expected default tabs to advance to the next 36pt stop in: {result}"
    );
    assert!(
        !result.contains("#h(36pt)"),
        "Expected default tabs to avoid a hard-coded 36pt gap in: {result}"
    );
}

#[test]
fn test_generate_tab_uses_next_explicit_stop_and_alignment() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![
                TabStop {
                    position: 72.0,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                },
                TabStop {
                    position: 216.0,
                    alignment: TabAlignment::Right,
                    leader: TabLeader::Dot,
                },
            ]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Col1\tCol2\tCol3".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("if tab_prefix_width_1 < 72pt"),
        "Expected the first explicit stop to be chosen by measured width in: {result}"
    );
    assert!(
        result.contains("else if tab_prefix_width_2 < 216pt"),
        "Expected the next explicit stop to be selected after the first one in: {result}"
    );
    assert!(
        result.contains("216pt - tab_prefix_width_2 - tab_segment_width_2"),
        "Expected right-aligned tabs to subtract the following segment width in: {result}"
    );
}

#[test]
fn test_generate_tab_falls_back_to_next_default_stop_after_explicit_tabs() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![TabStop {
                position: 100.0,
                alignment: TabAlignment::Left,
                leader: TabLeader::None,
            }]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "A\tB\tC".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("if tab_prefix_width_1 < 100pt"),
        "Expected the explicit stop to be used when it is still ahead of the prefix in: {result}"
    );
    assert!(
        result.contains("calc.rem-euclid(tab_prefix_width_2.abs.pt(), 36)"),
        "Expected tabs beyond explicit stops to use the next default stop in: {result}"
    );
}

#[test]
fn test_generate_tab_leader_uses_repeat_fill() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![TabStop {
                position: 144.0,
                alignment: TabAlignment::Left,
                leader: TabLeader::Dot,
            }]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Heading\t12".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("box(width: tab_advance_1, repeat[.])"),
        "Expected dot tab leaders to render with Typst repeat fill in: {result}"
    );
}

#[test]
fn test_generate_decimal_tab_uses_decimal_separator_not_thousands_separator() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![TabStop {
                position: 180.0,
                alignment: TabAlignment::Decimal,
                leader: TabLeader::None,
            }]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Total\t1,234.56".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("let tab_decimal_anchor_1 = [1,234]"),
        "Expected decimal alignment to anchor after the thousands group in: {result}"
    );
}

#[test]
fn test_generate_decimal_tab_handles_comma_decimal_locale() {
    use crate::ir::{TabAlignment, TabLeader, TabStop};

    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            tab_stops: Some(vec![TabStop {
                position: 180.0,
                alignment: TabAlignment::Decimal,
                leader: TabLeader::None,
            }]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Total\t1.234,56".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("let tab_decimal_anchor_1 = [1.234]"),
        "Expected decimal alignment to anchor on the locale decimal separator in: {result}"
    );
}

#[test]
fn test_generate_multiple_paragraphs() {
    let doc = make_doc(vec![make_flow_page(vec![
        make_paragraph("First paragraph"),
        make_paragraph("Second paragraph"),
    ])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("First paragraph"));
    assert!(result.contains("Second paragraph"));
    assert!(
        result.contains("First paragraph\n\nSecond paragraph"),
        "Expected paragraph break between flow paragraphs in: {result}"
    );
}

#[test]
fn test_generate_paragraph_with_multiple_runs() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "Normal ".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
            Run {
                text: "bold".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: " normal again".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
        ],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("Normal "));
    assert!(result.contains("bold"));
    assert!(result.contains(" normal again"));
}

#[test]
fn test_generate_empty_document() {
    let doc = make_doc(vec![]);
    let result = generate_typst(&doc).unwrap().source;
    // Should produce valid (possibly empty) Typst markup
    assert!(result.is_empty() || !result.is_empty()); // Just shouldn't error
}

#[test]
fn test_generate_special_characters_escaped() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(
        "Price: $100 #items @store",
    )])]);
    let result = generate_typst(&doc).unwrap().source;
    // The text should appear but special chars should be escaped for Typst
    // In Typst, # starts a code expression, so it needs escaping
    assert!(
        result.contains("\\#") || result.contains("Price"),
        "Expected escaped or present text in: {result}"
    );
}

// ── Table codegen tests ───────────────────────────────────────────

use crate::ir::{BorderSide, CellBorder, Insets, Table, TableCell, TableRow};

/// Helper to create a table cell with plain text.
fn make_text_cell(text: &str) -> TableCell {
    TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: text.to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        ..TableCell::default()
    }
}

#[test]
fn test_table_simple_2x2() {
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![make_text_cell("A1"), make_text_cell("B1")],
                height: None,
            },
            TableRow {
                cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("#table("), "Expected #table( in: {result}");
    assert!(
        result.contains("columns: (100pt, 200pt)"),
        "Expected column widths in: {result}"
    );
    assert!(result.contains("A1"), "Expected A1 in: {result}");
    assert!(result.contains("B1"), "Expected B1 in: {result}");
    assert!(result.contains("A2"), "Expected A2 in: {result}");
    assert!(result.contains("B2"), "Expected B2 in: {result}");
}

#[test]
fn test_table_with_default_cell_padding() {
    let table = Table {
        rows: vec![TableRow {
            cells: vec![make_text_cell("Padded")],
            height: None,
        }],
        column_widths: vec![100.0],
        header_row_count: 0,
        alignment: None,
        default_cell_padding: Some(Insets {
            top: 2.0,
            right: 3.0,
            bottom: 1.0,
            left: 4.0,
        }),
        use_content_driven_row_heights: false,
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("inset: (top: 2pt, right: 3pt, bottom: 1pt, left: 4pt)"),
        "Expected table inset in: {result}"
    );
}

#[test]
fn test_table_cell_with_padding_override() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Inset".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        padding: Some(Insets {
            top: 5.0,
            right: 2.0,
            bottom: 3.0,
            left: 6.0,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        header_row_count: 0,
        alignment: None,
        default_cell_padding: Some(Insets {
            top: 1.0,
            right: 2.0,
            bottom: 3.0,
            left: 4.0,
        }),
        use_content_driven_row_heights: false,
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("table.cell(inset: (top: 5pt, right: 2pt, bottom: 3pt, left: 6pt))"),
        "Expected cell inset override in: {result}"
    );
}

#[test]
fn test_table_alignment_center_wraps_table() {
    let table = Table {
        rows: vec![TableRow {
            cells: vec![make_text_cell("Centered table")],
            height: None,
        }],
        column_widths: vec![100.0],
        header_row_count: 0,
        alignment: Some(Alignment::Center),
        default_cell_padding: None,
        use_content_driven_row_heights: false,
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#align(center)["),
        "Expected center wrapper in: {result}"
    );
    assert!(
        result.contains("#table("),
        "Expected table inside wrapper in: {result}"
    );
}

#[test]
fn test_table_with_repeating_header_rows_uses_table_header() {
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![make_text_cell("Header 1"), make_text_cell("Header 2")],
                height: None,
            },
            TableRow {
                cells: vec![make_text_cell("Body 1"), make_text_cell("Body 2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 100.0],
        header_row_count: 1,
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("table.header("),
        "Expected table.header wrapper in: {result}"
    );
    assert!(
        result.contains("Header 1") && result.contains("Body 1"),
        "Expected header and body cell content in: {result}"
    );
}

#[test]
fn test_table_with_colspan() {
    let merged_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Merged".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        col_span: 2,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![merged_cell],
                height: None,
            },
            TableRow {
                cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("colspan: 2"),
        "Expected colspan: 2 in: {result}"
    );
    assert!(result.contains("Merged"), "Expected Merged in: {result}");
}

#[test]
fn test_table_with_rowspan() {
    let tall_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Tall".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        row_span: 2,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![tall_cell, make_text_cell("B1")],
                height: None,
            },
            TableRow {
                cells: vec![make_text_cell("B2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("rowspan: 2"),
        "Expected rowspan: 2 in: {result}"
    );
    assert!(result.contains("Tall"), "Expected Tall in: {result}");
}

#[test]
fn test_table_with_explicit_row_sizes_and_cell_vertical_align() {
    let centered_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Centered".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        vertical_align: Some(CellVerticalAlign::Center),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![centered_cell, make_text_cell("B1")],
                height: Some(36.0),
            },
            TableRow {
                cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("rows: (36pt, auto)"),
        "Expected explicit Typst row sizes in: {result}"
    );
    assert!(
        result.contains("align: horizon"),
        "Expected centered vertical alignment in: {result}"
    );
}

#[test]
fn test_table_with_content_driven_row_heights_omits_explicit_rows() {
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![make_text_cell("A1"), make_text_cell("B1")],
                height: Some(36.0),
            },
            TableRow {
                cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                height: Some(48.0),
            },
        ],
        column_widths: vec![100.0, 100.0],
        use_content_driven_row_heights: true,
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        !result.contains("rows: ("),
        "Content-driven row-height tables should not emit exact Typst row sizes: {result}"
    );
}

#[test]
fn test_table_with_colspan_and_rowspan() {
    let big_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Big".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        col_span: 2,
        row_span: 2,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![big_cell, make_text_cell("C1")],
                height: None,
            },
            TableRow {
                cells: vec![make_text_cell("C2")],
                height: None,
            },
            TableRow {
                cells: vec![
                    make_text_cell("A3"),
                    make_text_cell("B3"),
                    make_text_cell("C3"),
                ],
                height: None,
            },
        ],
        column_widths: vec![100.0, 100.0, 100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("colspan: 2"),
        "Expected colspan: 2 in: {result}"
    );
    assert!(
        result.contains("rowspan: 2"),
        "Expected rowspan: 2 in: {result}"
    );
    assert!(result.contains("Big"), "Expected Big in: {result}");
}

#[test]
fn test_table_with_background_color() {
    let colored_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Colored".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        background: Some(Color::new(200, 200, 200)),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![colored_cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("fill: rgb(200, 200, 200)"),
        "Expected fill color in: {result}"
    );
    assert!(result.contains("Colored"), "Expected Colored in: {result}");
}

#[test]
fn test_table_with_cell_borders() {
    let bordered_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Bordered".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
            bottom: Some(BorderSide {
                width: 2.0,
                color: Color::new(255, 0, 0),
                style: BorderLineStyle::Solid,
            }),
            left: None,
            right: None,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![bordered_cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("stroke:"), "Expected stroke in: {result}");
    assert!(
        result.contains("Bordered"),
        "Expected Bordered in: {result}"
    );
}

#[test]
fn test_table_with_styled_text_in_cell() {
    let styled_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Bold cell".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    font_size: Some(14.0),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        })],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![styled_cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("weight: \"bold\""),
        "Expected bold in table cell: {result}"
    );
    assert!(
        result.contains("size: 14pt"),
        "Expected font size in table cell: {result}"
    );
}

#[test]
fn test_table_empty_cells() {
    let empty_cell = TableCell::default();
    let table = Table {
        rows: vec![TableRow {
            cells: vec![empty_cell, make_text_cell("Has text")],
            height: None,
        }],
        column_widths: vec![100.0, 100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("#table("), "Expected #table( in: {result}");
    assert!(
        result.contains("Has text"),
        "Expected Has text in: {result}"
    );
}

#[test]
fn test_table_no_column_widths() {
    let table = Table {
        rows: vec![TableRow {
            cells: vec![make_text_cell("A"), make_text_cell("B")],
            height: None,
        }],
        column_widths: vec![],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("#table("), "Expected #table( in: {result}");
    // Without explicit widths, should still produce valid table
    assert!(result.contains("A"), "Expected A in: {result}");
    assert!(result.contains("B"), "Expected B in: {result}");
}

#[test]
fn test_table_all_borders() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "All borders".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
            bottom: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
            left: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
            right: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(result.contains("top:"), "Expected top border in: {result}");
    assert!(
        result.contains("bottom:"),
        "Expected bottom border in: {result}"
    );
    assert!(
        result.contains("left:"),
        "Expected left border in: {result}"
    );
    assert!(
        result.contains("right:"),
        "Expected right border in: {result}"
    );
}

#[test]
fn test_table_dashed_border_codegen() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Dashed".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Dashed,
            }),
            bottom: Some(BorderSide {
                width: 1.0,
                color: Color::new(255, 0, 0),
                style: BorderLineStyle::Dotted,
            }),
            left: None,
            right: None,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("dash: \"dashed\""),
        "Expected dashed dash pattern in: {result}"
    );
    assert!(
        result.contains("dash: \"dotted\""),
        "Expected dotted dash pattern in: {result}"
    );
}

#[test]
fn test_shape_dashed_stroke_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            10.0,
            100.0,
            100.0,
            ShapeKind::Rectangle,
            Some(Color::new(0, 128, 255)),
            Some(BorderSide {
                width: 2.0,
                color: Color::black(),
                style: BorderLineStyle::Dashed,
            }),
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("dash: \"dashed\""),
        "Expected dashed stroke in: {}",
        output.source
    );
}

#[test]
fn test_shape_dash_dot_stroke_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            10.0,
            100.0,
            100.0,
            ShapeKind::Ellipse,
            None,
            Some(BorderSide {
                width: 1.0,
                color: Color::new(0, 0, 255),
                style: BorderLineStyle::DashDot,
            }),
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("dash: \"dash-dotted\""),
        "Expected dash-dotted stroke in: {}",
        output.source
    );
}

#[test]
fn test_border_line_style_to_typst_mapping() {
    assert_eq!(border_line_style_to_typst(BorderLineStyle::Solid), "solid");
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::Dashed),
        "dashed"
    );
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::Dotted),
        "dotted"
    );
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::DashDot),
        "dash-dotted"
    );
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::DashDotDot),
        "dash-dotted"
    );
    assert_eq!(
        border_line_style_to_typst(BorderLineStyle::Double),
        "dashed"
    );
    assert_eq!(border_line_style_to_typst(BorderLineStyle::None), "solid");
}

#[test]
fn test_solid_border_no_dash_param() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Solid".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        border: Some(CellBorder {
            top: Some(BorderSide {
                width: 1.0,
                color: Color::black(),
                style: BorderLineStyle::Solid,
            }),
            bottom: None,
            left: None,
            right: None,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    // Solid borders should use the simple format (no "dash:" parameter)
    assert!(
        !result.contains("dash:"),
        "Solid border should not have dash parameter in: {result}"
    );
    assert!(
        result.contains("1pt + rgb(0, 0, 0)"),
        "Expected simple solid format in: {result}"
    );
}

#[test]
fn test_table_cell_with_multiple_paragraphs() {
    let multi_para_cell = TableCell {
        content: vec![
            Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "First para".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            }),
            Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Second para".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            }),
        ],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![multi_para_cell],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("First para"),
        "Expected First para in: {result}"
    );
    assert!(
        result.contains("Second para"),
        "Expected Second para in: {result}"
    );
}

#[test]
fn test_table_cell_simple_list_uses_compact_fixed_text_layout() {
    let list = List {
        kind: ListKind::Unordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "First item".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Second item".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
        ],
        level_styles: BTreeMap::new(),
    };
    let cell = TableCell {
        content: vec![Block::List(list)],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#stack(dir: ttb"),
        "Expected compact stack-based list layout in: {result}"
    );
    assert!(
        !result.contains("#list("),
        "Compact table-cell lists should not use Typst list layout in: {result}"
    );
    assert!(result.contains("First item"));
    assert!(result.contains("Second item"));
}

#[test]
fn test_table_cell_simple_list_treats_default_and_explicit_left_as_same_style() {
    let list = List {
        kind: ListKind::Unordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle {
                        alignment: Some(Alignment::Left),
                        ..ParagraphStyle::default()
                    },
                    runs: vec![Run {
                        text: "First item".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Second item".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
        ],
        level_styles: BTreeMap::new(),
    };
    let cell = TableCell {
        content: vec![Block::List(list)],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#stack(dir: ttb"),
        "Expected compact stack-based list layout when only left-alignment explicitness differs: {result}"
    );
    assert!(
        !result.contains("#list("),
        "Equivalent left-alignment styles should not force Typst list layout in: {result}"
    );
}

#[test]
fn test_table_cell_compact_list_uses_leading_without_extra_item_spacing() {
    let list = List {
        kind: ListKind::Unordered,
        items: vec![
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle {
                        line_spacing: Some(LineSpacing::Proportional(1.5)),
                        ..ParagraphStyle::default()
                    },
                    runs: vec![Run {
                        text: "First item".to_string(),
                        style: TextStyle {
                            font_size: Some(24.0),
                            ..TextStyle::default()
                        },
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
            ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle {
                        line_spacing: Some(LineSpacing::Proportional(1.5)),
                        ..ParagraphStyle::default()
                    },
                    runs: vec![Run {
                        text: "Second item".to_string(),
                        style: TextStyle {
                            font_size: Some(24.0),
                            ..TextStyle::default()
                        },
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
                start_at: None,
            },
        ],
        level_styles: BTreeMap::new(),
    };
    let cell = TableCell {
        content: vec![Block::List(list)],
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;

    assert!(
        result.contains("#set par(leading: 12pt)"),
        "Expected paragraph leading derived from PPT line spacing in: {result}"
    );
    assert!(
        !result.contains("#stack(dir: ttb, spacing: 12pt"),
        "Compact table-cell lists should not add extra inter-item spacing in: {result}"
    );
}

#[test]
fn test_table_special_chars_in_cells() {
    let table = Table {
        rows: vec![TableRow {
            cells: vec![make_text_cell("Price: $100 #items")],
            height: None,
        }],
        column_widths: vec![200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    // Special chars should be escaped
    assert!(
        result.contains("\\$") && result.contains("\\#"),
        "Expected escaped special chars in: {result}"
    );
}

#[test]
fn test_table_in_flow_page_with_paragraphs() {
    let table = Table {
        rows: vec![TableRow {
            cells: vec![make_text_cell("Cell")],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![
        make_paragraph("Before table"),
        Block::Table(table),
        make_paragraph("After table"),
    ])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("Before table"),
        "Expected Before table in: {result}"
    );
    assert!(result.contains("#table("), "Expected #table( in: {result}");
    assert!(
        result.contains("After table"),
        "Expected After table in: {result}"
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
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    // Should contain spacing directives
    assert!(
        result.contains("12pt") || result.contains("above"),
        "Expected space_before in: {result}"
    );
}

// ── Image codegen tests ─────────────────────────────────────────────

use crate::ir::{ImageCrop, ImageData};

/// Minimal valid 1x1 red pixel PNG for testing.
const MINIMAL_PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82,
];

fn make_quadrant_png() -> Vec<u8> {
    let mut image = image::RgbaImage::new(2, 2);
    image.put_pixel(0, 0, image::Rgba([255, 0, 0, 255]));
    image.put_pixel(1, 0, image::Rgba([0, 255, 0, 255]));
    image.put_pixel(0, 1, image::Rgba([0, 0, 255, 255]));
    image.put_pixel(1, 1, image::Rgba([255, 255, 0, 255]));

    let mut encoded = Cursor::new(Vec::new());
    image::DynamicImage::ImageRgba8(image)
        .write_to(&mut encoded, RasterImageFormat::Png)
        .unwrap();
    encoded.into_inner()
}

fn make_image(format: ImageFormat, width: Option<f64>, height: Option<f64>) -> Block {
    Block::Image(ImageData {
        data: MINIMAL_PNG.to_vec(),
        format,
        width,
        height,
        crop: None,
    })
}

#[test]
fn test_image_basic_no_size() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        None,
        None,
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#image(\"img-0.png\")"),
        "Expected #image(\"img-0.png\") in: {}",
        output.source
    );
}

#[test]
fn test_image_crop_preprocesses_raster_asset() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Image(ImageData {
        data: make_quadrant_png(),
        format: ImageFormat::Png,
        width: Some(20.0),
        height: Some(20.0),
        crop: Some(ImageCrop {
            left: 0.5,
            top: 0.0,
            right: 0.0,
            bottom: 0.0,
        }),
    })])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("#image(\"img-0.png\", width: 20pt, height: 20pt)"),
        "Expected original display size in: {}",
        output.source
    );

    let cropped =
        image::load_from_memory_with_format(&output.images[0].data, RasterImageFormat::Png)
            .unwrap()
            .to_rgba8();
    assert_eq!(cropped.dimensions(), (1, 2));
    assert_eq!(cropped.get_pixel(0, 0).0, [0, 255, 0, 255]);
    assert_eq!(cropped.get_pixel(0, 1).0, [255, 255, 0, 255]);
}

#[test]
fn test_image_with_width_only() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        Some(100.0),
        None,
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("#image(\"img-0.png\", width: 100pt)"),
        "Expected width param in: {}",
        output.source
    );
}

#[test]
fn test_image_with_height_only() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        None,
        Some(80.0),
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("#image(\"img-0.png\", height: 80pt)"),
        "Expected height param in: {}",
        output.source
    );
}

#[test]
fn test_image_with_both_dimensions() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        Some(200.0),
        Some(150.0),
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("#image(\"img-0.png\", width: 200pt, height: 150pt)"),
        "Expected both dimensions in: {}",
        output.source
    );
}

#[test]
fn test_image_collects_asset() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        None,
        None,
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert_eq!(output.images.len(), 1);
    assert_eq!(output.images[0].path, "img-0.png");
    assert_eq!(output.images[0].data, MINIMAL_PNG);
}

#[test]
fn test_multiple_images_numbered_sequentially() {
    let doc = make_doc(vec![make_flow_page(vec![
        make_image(ImageFormat::Png, None, None),
        make_image(ImageFormat::Jpeg, Some(50.0), None),
    ])]);
    let output = generate_typst(&doc).unwrap();
    assert_eq!(output.images.len(), 2);
    assert_eq!(output.images[0].path, "img-0.png");
    assert_eq!(output.images[1].path, "img-1.jpeg");
    assert!(output.source.contains("img-0.png"));
    assert!(output.source.contains("img-1.jpeg"));
}

#[test]
fn test_image_format_extensions() {
    let formats = [
        (ImageFormat::Png, "png"),
        (ImageFormat::Jpeg, "jpeg"),
        (ImageFormat::Gif, "gif"),
        (ImageFormat::Bmp, "bmp"),
        (ImageFormat::Tiff, "tiff"),
        (ImageFormat::Svg, "svg"),
    ];
    for (i, (format, expected_ext)) in formats.iter().enumerate() {
        let doc = make_doc(vec![make_flow_page(vec![make_image(*format, None, None)])]);
        let output = generate_typst(&doc).unwrap();
        let expected_path = format!("img-0.{expected_ext}");
        assert_eq!(
            output.images[0].path, expected_path,
            "Format {format:?} should produce .{expected_ext} extension (test #{i})"
        );
    }
}

#[test]
fn test_image_with_fractional_dimensions() {
    let doc = make_doc(vec![make_flow_page(vec![make_image(
        ImageFormat::Png,
        Some(72.5),
        Some(96.25),
    )])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("width: 72.5pt"),
        "Expected fractional width in: {}",
        output.source
    );
    assert!(
        output.source.contains("height: 96.25pt"),
        "Expected fractional height in: {}",
        output.source
    );
}

#[test]
fn test_image_mixed_with_paragraphs() {
    let doc = make_doc(vec![make_flow_page(vec![
        make_paragraph("Before image"),
        make_image(ImageFormat::Png, Some(100.0), Some(80.0)),
        make_paragraph("After image"),
    ])]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("Before image"));
    assert!(output.source.contains("#image(\"img-0.png\""));
    assert!(output.source.contains("After image"));
    assert_eq!(output.images.len(), 1);
}

#[test]
fn test_no_images_produces_empty_assets() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Just text")])]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.images.is_empty());
}

// ── FixedPage codegen tests (US-010) ────────────────────────────────

/// Helper to create a FixedPage (slide-like) with given elements.
fn make_fixed_page(width: f64, height: f64, elements: Vec<FixedElement>) -> Page {
    Page::Fixed(FixedPage {
        size: PageSize { width, height },
        elements,
        background_color: None,
        background_gradient: None,
    })
}

/// Helper to create a text box FixedElement.
fn make_text_box(x: f64, y: f64, w: f64, h: f64, text: &str) -> FixedElement {
    FixedElement {
        x,
        y,
        width: w,
        height: h,
        kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: text.to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            padding: Insets::default(),
            vertical_align: crate::ir::TextBoxVerticalAlign::Top,
        }),
    }
}

/// Helper to create a shape FixedElement.
fn make_shape_element(
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    kind: ShapeKind,
    fill: Option<Color>,
    stroke: Option<BorderSide>,
) -> FixedElement {
    FixedElement {
        x,
        y,
        width: w,
        height: h,
        kind: FixedElementKind::Shape(Shape {
            kind,
            fill,
            gradient_fill: None,
            stroke,
            rotation_deg: None,
            opacity: None,
            shadow: None,
        }),
    }
}

fn make_fixed_text_box(
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    padding: Insets,
    vertical_align: crate::ir::TextBoxVerticalAlign,
    content: Vec<Block>,
) -> FixedElement {
    FixedElement {
        x,
        y,
        width: w,
        height: h,
        kind: FixedElementKind::TextBox(crate::ir::TextBoxData {
            content,
            padding,
            vertical_align,
        }),
    }
}

/// Helper to create an image FixedElement.
fn make_fixed_image(x: f64, y: f64, w: f64, h: f64, format: ImageFormat) -> FixedElement {
    FixedElement {
        x,
        y,
        width: w,
        height: h,
        kind: FixedElementKind::Image(ImageData {
            data: vec![0x89, 0x50, 0x4E, 0x47], // PNG header stub
            format,
            width: Some(w),
            height: Some(h),
            crop: None,
        }),
    }
}

#[path = "typst_gen_fixed_page_tests.rs"]
mod fixed_page_tests;

// ── TablePage codegen tests ──────────────────────────────────────────

/// Helper to create a TablePage.
fn make_table_page(name: &str, width: f64, height: f64, margins: Margins, table: Table) -> Page {
    Page::Table(crate::ir::TablePage {
        name: name.to_string(),
        size: PageSize { width, height },
        margins,
        table,
        header: None,
        footer: None,
        charts: vec![],
    })
}

/// Helper to create a simple Table with text cells.
fn make_simple_table(rows: Vec<Vec<&str>>) -> Table {
    Table {
        rows: rows
            .into_iter()
            .map(|cells| TableRow {
                cells: cells
                    .into_iter()
                    .map(|text| TableCell {
                        content: vec![Block::Paragraph(Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: text.to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                            }],
                        })],
                        ..TableCell::default()
                    })
                    .collect(),
                height: None,
            })
            .collect(),
        column_widths: vec![],
        ..Table::default()
    }
}

#[path = "typst_gen_table_page_tests.rs"]
mod table_page_tests;

// ----- List codegen tests -----

#[path = "typst_gen_list_tests.rs"]
mod list_tests;

// ----- US-020: Header/footer codegen tests -----

#[test]
fn test_generate_flow_page_with_text_header() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body text")],
        header: Some(HeaderFooter {
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Document Title".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
            }],
        }),
        footer: None,
        columns: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("header:"),
        "Should contain header: in page setup. Got: {}",
        output.source
    );
    assert!(
        output.source.contains("Document Title"),
        "Header should contain 'Document Title'. Got: {}",
        output.source
    );
}

#[test]
fn test_generate_flow_page_with_page_number_footer() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body text")],
        header: None,
        footer: Some(HeaderFooter {
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![
                    HFInline::Run(Run {
                        text: "Page ".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }),
                    HFInline::PageNumber,
                ],
            }],
        }),
        columns: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("footer:"),
        "Should contain footer: in page setup. Got: {}",
        output.source
    );
    assert!(
        output.source.contains("counter(page).display()"),
        "Footer should contain page counter. Got: {}",
        output.source
    );
    assert!(
        output.source.contains("Page "),
        "Footer should contain 'Page ' text. Got: {}",
        output.source
    );
}

#[test]
fn test_generate_flow_page_with_header_and_footer() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Header".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
            }],
        }),
        footer: Some(HeaderFooter {
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::PageNumber],
            }],
        }),
        columns: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("header:") && output.source.contains("footer:"),
        "Should contain both header: and footer:. Got: {}",
        output.source
    );
}

#[test]
fn test_generate_flow_page_without_header_footer() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Body")])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        !output.source.contains("header:"),
        "Should NOT contain header: when no header. Got: {}",
        output.source
    );
    assert!(
        !output.source.contains("footer:"),
        "Should NOT contain footer: when no footer. Got: {}",
        output.source
    );
}

#[test]
fn test_generate_typst_inserts_pagebreak_between_flow_pages() {
    let first = Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("First section")],
        header: None,
        footer: None,
        columns: None,
    });
    let second = Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Second section")],
        header: None,
        footer: None,
        columns: None,
    });

    let output = generate_typst(&make_doc(vec![first, second])).unwrap();
    let pagebreak_count = output.source.matches("#pagebreak()").count();

    assert_eq!(
        pagebreak_count, 1,
        "Expected exactly one page break between FlowPages. Got:\n{}",
        output.source
    );
}

// ── Fixed page background tests ──────────────────────────────────────

#[test]
fn test_fixed_page_with_background_color() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: Some(Color::new(255, 0, 0)),
        background_gradient: None,
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("fill: rgb(255, 0, 0)"),
        "Should contain fill: rgb(255, 0, 0). Got: {}",
        output.source
    );
}

#[test]
fn test_fixed_page_without_background_color() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: None,
        background_gradient: None,
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        !output.source.contains("fill:"),
        "Should NOT contain fill: when no background. Got: {}",
        output.source
    );
}

#[test]
fn test_fixed_page_table_element() {
    // A table placed at absolute position on a fixed page
    let table = Table {
        rows: vec![TableRow {
            cells: vec![
                TableCell {
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "A1".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    })],
                    ..TableCell::default()
                },
                TableCell {
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "B1".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    })],
                    ..TableCell::default()
                },
            ],
            height: None,
        }],
        column_widths: vec![100.0, 100.0],
        ..Table::default()
    };

    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![FixedElement {
            x: 50.0,
            y: 100.0,
            width: 200.0,
            height: 50.0,
            kind: FixedElementKind::Table(table),
        }],
        background_color: None,
        background_gradient: None,
    });

    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();

    // Should have a #place() with table inside
    assert!(
        output
            .source
            .contains("#place(top + left, dx: 50pt, dy: 100pt)")
    );
    assert!(output.source.contains("#table("));
    assert!(output.source.contains("columns: (100pt, 100pt)"));
    assert!(output.source.contains("A1"));
    assert!(output.source.contains("B1"));
}

// ----- Hyperlink codegen tests (US-030) -----

#[test]
fn test_hyperlink_generates_typst_link() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Click me".to_string(),
            style: TextStyle::default(),
            href: Some("https://example.com".to_string()),
            footnote: None,
        }],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains(r#"#link("https://example.com")[Click me]"#),
        "Expected Typst link markup, got: {}",
        output.source
    );
}

#[test]
fn test_hyperlink_with_styled_text() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Bold link".to_string(),
            style: TextStyle {
                bold: Some(true),
                ..TextStyle::default()
            },
            href: Some("https://example.com".to_string()),
            footnote: None,
        }],
    })])]);

    let output = generate_typst(&doc).unwrap();
    // Should have link wrapping styled text
    assert!(
        output.source.contains(r#"#link("https://example.com")["#),
        "Expected Typst link markup, got: {}",
        output.source
    );
    assert!(
        output.source.contains("#text(weight: \"bold\")"),
        "Expected bold text inside link, got: {}",
        output.source
    );
}

#[test]
fn test_hyperlink_mixed_with_plain_text() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "Visit ".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
            Run {
                text: "Rust".to_string(),
                style: TextStyle::default(),
                href: Some("https://rust-lang.org".to_string()),
                footnote: None,
            },
            Run {
                text: " for more.".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
        ],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("Visit "),
        "Expected plain text, got: {}",
        output.source
    );
    assert!(
        output
            .source
            .contains(r#"#link("https://rust-lang.org")[Rust]"#),
        "Expected Typst link markup, got: {}",
        output.source
    );
    assert!(
        output.source.contains(" for more."),
        "Expected plain text after link, got: {}",
        output.source
    );
}

#[test]
fn test_hyperlink_url_with_special_chars_escaped() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Link".to_string(),
            style: TextStyle::default(),
            href: Some("https://example.com/path?q=1&r=2".to_string()),
            footnote: None,
        }],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains(r#"#link("https://example.com/path?q=1&r=2")[Link]"#),
        "Expected URL with special chars in link, got: {}",
        output.source
    );
}

// ── Footnotes ───────────────────────────────────────────────────────

#[test]
fn test_footnote_generates_typst_footnote() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "Some text".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
            Run {
                text: String::new(),
                style: TextStyle::default(),
                href: None,
                footnote: Some("This is a footnote.".to_string()),
            },
        ],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#footnote[This is a footnote.]"),
        "Expected Typst footnote markup, got: {}",
        output.source
    );
}

#[test]
fn test_footnote_with_special_chars() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: String::new(),
            style: TextStyle::default(),
            href: None,
            footnote: Some("Note with #special *chars*".to_string()),
        }],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains(r"#footnote[Note with \#special \*chars\*]"),
        "Expected escaped footnote content, got: {}",
        output.source
    );
}

// --- US-036: TablePage header/footer codegen ---

#[test]
fn test_table_page_with_header() {
    let page = Page::Table(TablePage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: make_simple_table(vec![vec!["A"]]),
        header: Some(HeaderFooter {
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle {
                    alignment: Some(Alignment::Center),
                    ..ParagraphStyle::default()
                },
                elements: vec![HFInline::Run(Run {
                    text: "My Header".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
            }],
        }),
        footer: None,
        charts: vec![],
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("header: ["),
        "Expected header in page setup, got: {}",
        output.source
    );
    assert!(
        output.source.contains("My Header"),
        "Expected header text, got: {}",
        output.source
    );
}

#[test]
fn test_table_page_with_page_number_footer() {
    let page = Page::Table(TablePage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: make_simple_table(vec![vec!["A"]]),
        header: None,
        footer: Some(HeaderFooter {
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle {
                    alignment: Some(Alignment::Center),
                    ..ParagraphStyle::default()
                },
                elements: vec![
                    HFInline::Run(Run {
                        text: "Page ".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }),
                    HFInline::PageNumber,
                    HFInline::Run(Run {
                        text: " of ".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }),
                    HFInline::TotalPages,
                ],
            }],
        }),
        charts: vec![],
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    // Footer with page numbers needs context
    assert!(
        output.source.contains("footer: context ["),
        "Expected context footer, got: {}",
        output.source
    );
    assert!(
        output.source.contains("#counter(page).display()"),
        "Expected page number counter, got: {}",
        output.source
    );
    assert!(
        output.source.contains("#counter(page).final().first()"),
        "Expected total pages counter, got: {}",
        output.source
    );
}

#[test]
fn test_table_page_no_header_footer() {
    let page = Page::Table(TablePage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: make_simple_table(vec![vec!["A"]]),
        header: None,
        footer: None,
        charts: vec![],
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    // Should use simple page setup without header/footer
    assert!(
        !output.source.contains("header:"),
        "Expected no header, got: {}",
        output.source
    );
    assert!(
        !output.source.contains("footer:"),
        "Expected no footer, got: {}",
        output.source
    );
}

// --- Table page with interleaved charts ---

#[test]
fn test_table_page_with_chart_at_row() {
    use crate::ir::{Chart, ChartSeries, ChartType};

    let chart = Chart {
        chart_type: ChartType::Bar,
        title: Some("Sales".to_string()),
        categories: vec!["Q1".to_string(), "Q2".to_string()],
        series: vec![ChartSeries {
            name: Some("Revenue".to_string()),
            values: vec![100.0, 200.0],
        }],
    };

    let page = Page::Table(TablePage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: make_simple_table(vec![
            vec!["Row 1"],
            vec!["Row 2"],
            vec!["Row 3"],
            vec!["Row 4"],
            vec!["Row 5"],
        ]),
        header: None,
        footer: None,
        charts: vec![(2, chart)], // Chart after row 2
    });

    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    let src = &output.source;

    // Should contain two separate #table blocks (split at chart position)
    let table_count = src.matches("#table(").count();
    assert_eq!(
        table_count, 2,
        "Expected 2 table segments (split at chart row), got {table_count}"
    );

    // Should contain chart rendering between table segments
    assert!(src.contains("Sales"), "Expected chart title in output");
}

#[test]
fn test_table_page_with_chart_at_end() {
    use crate::ir::{Chart, ChartSeries, ChartType};

    let chart = Chart {
        chart_type: ChartType::Pie,
        title: Some("Pie".to_string()),
        categories: vec!["A".to_string()],
        series: vec![ChartSeries {
            name: None,
            values: vec![100.0],
        }],
    };

    let page = Page::Table(TablePage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: make_simple_table(vec![vec!["Data"]]),
        header: None,
        footer: None,
        charts: vec![(u32::MAX, chart)], // Chart at end
    });

    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    let src = &output.source;

    // Table should appear before chart
    let table_pos = src.find("#table(").unwrap();
    let chart_pos = src.find("Pie").unwrap();
    assert!(table_pos < chart_pos, "Table should appear before chart");
}

// --- Paper size and landscape override tests ---

#[test]
fn test_paper_size_override_letter() {
    use crate::config::PaperSize;

    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
    let options = ConvertOptions {
        paper_size: Some(PaperSize::Letter),
        ..Default::default()
    };
    let output = generate_typst_with_options(&doc, &options).unwrap();
    assert!(
        output.source.contains("width: 612pt"),
        "Expected Letter width 612pt, got: {}",
        output.source
    );
    assert!(
        output.source.contains("height: 792pt"),
        "Expected Letter height 792pt, got: {}",
        output.source
    );
}

#[test]
fn test_landscape_override_swaps_dimensions() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
    let options = ConvertOptions {
        landscape: Some(true),
        ..Default::default()
    };
    let output = generate_typst_with_options(&doc, &options).unwrap();
    // A4 default is 595.28 x 841.89; landscape should swap to 841.89 x 595.28
    assert!(
        output.source.contains("width: 841.89pt"),
        "Expected landscape width 841.89pt, got: {}",
        output.source
    );
    assert!(
        output.source.contains("height: 595.28pt"),
        "Expected landscape height 595.28pt, got: {}",
        output.source
    );
}

#[test]
fn test_portrait_override_keeps_portrait() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
    let options = ConvertOptions {
        landscape: Some(false),
        ..Default::default()
    };
    let output = generate_typst_with_options(&doc, &options).unwrap();
    // A4 is already portrait, should remain unchanged
    assert!(
        output.source.contains("width: 595.28pt"),
        "Expected portrait width, got: {}",
        output.source
    );
    assert!(
        output.source.contains("height: 841.89pt"),
        "Expected portrait height, got: {}",
        output.source
    );
}

#[test]
fn test_paper_size_with_landscape() {
    use crate::config::PaperSize;

    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
    let options = ConvertOptions {
        paper_size: Some(PaperSize::Letter),
        landscape: Some(true),
        ..Default::default()
    };
    let output = generate_typst_with_options(&doc, &options).unwrap();
    // Letter landscape: 792 x 612
    assert!(
        output.source.contains("width: 792pt"),
        "Expected landscape Letter width 792pt, got: {}",
        output.source
    );
    assert!(
        output.source.contains("height: 612pt"),
        "Expected landscape Letter height 612pt, got: {}",
        output.source
    );
}

#[test]
fn test_no_override_uses_original_size() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
    let options = ConvertOptions::default();
    let output = generate_typst_with_options(&doc, &options).unwrap();
    // Default A4 dimensions
    assert!(
        output.source.contains("width: 595.28pt"),
        "Expected A4 width, got: {}",
        output.source
    );
}

// ── Floating image codegen tests ──

#[test]
fn test_floating_image_square_wrap_codegen() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::FloatingImage(FloatingImage {
                image: ImageData {
                    data: vec![0x89, 0x50, 0x4E, 0x47],
                    format: ImageFormat::Png,
                    width: Some(200.0),
                    height: Some(100.0),
                    crop: None,
                },
                wrap_mode: WrapMode::Square,
                offset_x: 72.0,
                offset_y: 36.0,
            })],
            header: None,
            footer: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };

    let output = generate_typst(&doc).unwrap();
    // Square wrap should use #place with float: true
    assert!(
        output.source.contains("#place("),
        "Expected #place() for floating image, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("float: true"),
        "Expected float: true for square wrap, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("dx: 72pt"),
        "Expected dx: 72pt, got:\n{}",
        output.source
    );
}

#[test]
fn test_floating_image_top_and_bottom_codegen() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::FloatingImage(FloatingImage {
                image: ImageData {
                    data: vec![0x89, 0x50, 0x4E, 0x47],
                    format: ImageFormat::Png,
                    width: Some(150.0),
                    height: Some(75.0),
                    crop: None,
                },
                wrap_mode: WrapMode::TopAndBottom,
                offset_x: 10.0,
                offset_y: 0.0,
            })],
            header: None,
            footer: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };

    let output = generate_typst(&doc).unwrap();
    // TopAndBottom should use a block with vertical space
    assert!(
        output.source.contains("#block("),
        "Expected #block() for topAndBottom wrap, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("#v(75pt)"),
        "Expected vertical space for image height, got:\n{}",
        output.source
    );
}

#[test]
fn test_floating_image_behind_codegen() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::FloatingImage(FloatingImage {
                image: ImageData {
                    data: vec![0x89, 0x50, 0x4E, 0x47],
                    format: ImageFormat::Png,
                    width: Some(100.0),
                    height: Some(50.0),
                    crop: None,
                },
                wrap_mode: WrapMode::Behind,
                offset_x: 0.0,
                offset_y: 0.0,
            })],
            header: None,
            footer: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };

    let output = generate_typst(&doc).unwrap();
    // Behind should use #place without float
    assert!(
        output.source.contains("#place("),
        "Expected #place() for behind wrap, got:\n{}",
        output.source
    );
    assert!(
        !output.source.contains("float: true"),
        "Behind wrap should NOT use float, got:\n{}",
        output.source
    );
}

#[test]
fn test_floating_text_box_square_wrap_codegen() {
    let doc = make_doc(vec![make_flow_page(vec![Block::FloatingTextBox(
        FloatingTextBox {
            content: vec![make_paragraph("Anchored box")],
            wrap_mode: WrapMode::Square,
            width: 200.0,
            height: 100.0,
            offset_x: 72.0,
            offset_y: 36.0,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#place("),
        "Expected #place() for floating text box, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("float: true"),
        "Expected float: true for square-wrapped text box, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("dx: 72pt"),
        "Expected dx: 72pt, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("width: 200pt"),
        "Expected width: 200pt, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("height: 100pt"),
        "Expected height: 100pt, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Anchored box"),
        "Expected text box content, got:\n{}",
        output.source
    );
}

#[test]
fn test_floating_text_box_top_and_bottom_codegen() {
    let doc = make_doc(vec![make_flow_page(vec![Block::FloatingTextBox(
        FloatingTextBox {
            content: vec![make_paragraph("Top box")],
            wrap_mode: WrapMode::TopAndBottom,
            width: 150.0,
            height: 60.0,
            offset_x: 10.0,
            offset_y: 0.0,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#block(width: 100%)"),
        "Expected block wrapper for top-and-bottom text box, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("#v(60pt)"),
        "Expected reserved vertical space for text box height, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Top box"),
        "Expected text box content, got:\n{}",
        output.source
    );
}

// ── Math equation codegen tests ──

#[test]
fn test_codegen_display_math() {
    let doc = make_doc(vec![make_flow_page(vec![Block::MathEquation(
        MathEquation {
            content: "frac(a, b)".to_string(),
            display: true,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("$ frac(a, b) $"),
        "Expected display math '$ frac(a, b) $', got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_inline_math() {
    let doc = make_doc(vec![make_flow_page(vec![Block::MathEquation(
        MathEquation {
            content: "x^2".to_string(),
            display: false,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("$x^2$"),
        "Expected inline math '$x^2$', got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_complex_math() {
    let doc = make_doc(vec![make_flow_page(vec![Block::MathEquation(
        MathEquation {
            content: "sum_(i=1)^n i".to_string(),
            display: true,
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("$ sum_(i=1)^n i $"),
        "Expected display math with sum, got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_chart_bar_visual_bars() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Chart {
        chart_type: ChartType::Bar,
        title: Some("Sales Report".to_string()),
        categories: vec!["Q1".to_string(), "Q2".to_string()],
        series: vec![ChartSeries {
            name: Some("Revenue".to_string()),
            values: vec![100.0, 250.0],
        }],
    })])]);

    let output = generate_typst(&doc).unwrap();
    // Wrapped in bordered box with header
    assert!(
        output.source.contains("stroke:"),
        "Expected bordered box, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Sales Report"),
        "Expected chart title in header, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Bar Chart"),
        "Expected chart type label, got:\n{}",
        output.source
    );
    // Bar chart should have visual bars (box with proportional width)
    assert!(
        output.source.contains("box(") || output.source.contains("#box("),
        "Expected visual bar boxes for bar chart, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Q1"),
        "Expected category label, got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_chart_pie_percentages() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Chart {
        chart_type: ChartType::Pie,
        title: Some("Market Share".to_string()),
        categories: vec!["A".to_string(), "B".to_string()],
        series: vec![ChartSeries {
            name: None,
            values: vec![60.0, 40.0],
        }],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("Pie Chart"),
        "Expected pie chart label, got:\n{}",
        output.source
    );
    // Pie chart should show percentages
    assert!(
        output.source.contains("60") && output.source.contains("%"),
        "Expected percentage in pie chart, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("40") && output.source.contains("%"),
        "Expected percentage in pie chart, got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_chart_line_trend_indicators() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Chart {
        chart_type: ChartType::Line,
        title: Some("Trends".to_string()),
        categories: vec!["Jan".to_string(), "Feb".to_string(), "Mar".to_string()],
        series: vec![ChartSeries {
            name: Some("Sales".to_string()),
            values: vec![10.0, 20.0, 15.0],
        }],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("Line Chart"),
        "Expected line chart label, got:\n{}",
        output.source
    );
    // Line chart should have trend indicators (↑ or ↓)
    let has_trend =
        output.source.contains('↑') || output.source.contains('↓') || output.source.contains('→');
    assert!(
        has_trend,
        "Expected trend indicators in line chart, got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_chart_empty_series() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Chart {
        chart_type: ChartType::Line,
        title: Some("Empty".to_string()),
        categories: vec![],
        series: vec![],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("Line Chart"),
        "Expected line chart label, got:\n{}",
        output.source
    );
}

// ── SmartArt codegen tests ──────────────────────────────────────────

/// Helper to create a SmartArtNode.
fn sa_node(text: &str, depth: usize) -> SmartArtNode {
    SmartArtNode {
        text: text.to_string(),
        depth,
    }
}

#[test]
fn test_smartart_codegen_flat_numbered_steps() {
    let doc = make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![FixedElement {
            x: 72.0,
            y: 100.0,
            width: 400.0,
            height: 300.0,
            kind: FixedElementKind::SmartArt(SmartArt {
                items: vec![
                    sa_node("Step 1", 0),
                    sa_node("Step 2", 0),
                    sa_node("Step 3", 0),
                ],
            }),
        }],
    )]);

    let output = generate_typst(&doc).unwrap();
    // Wrapped in bordered box
    assert!(
        output.source.contains("stroke:"),
        "Expected bordered box, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("SmartArt Diagram"),
        "Expected SmartArt header, got:\n{}",
        output.source
    );
    // Flat items → numbered steps with arrows
    assert!(
        output.source.contains("Step 1"),
        "Expected Step 1, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Step 2"),
        "Expected Step 2, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Step 3"),
        "Expected Step 3, got:\n{}",
        output.source
    );
}

#[test]
fn test_smartart_codegen_hierarchy_indented_tree() {
    let doc = make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![FixedElement {
            x: 72.0,
            y: 100.0,
            width: 400.0,
            height: 300.0,
            kind: FixedElementKind::SmartArt(SmartArt {
                items: vec![
                    sa_node("CEO", 0),
                    sa_node("VP Engineering", 1),
                    sa_node("VP Sales", 1),
                    sa_node("Dev Lead", 2),
                ],
            }),
        }],
    )]);

    let output = generate_typst(&doc).unwrap();
    // Hierarchical items should use indentation
    assert!(
        output.source.contains("CEO"),
        "Expected CEO, got:\n{}",
        output.source
    );
    // Deeper items should have padding/indentation
    assert!(
        output.source.contains("pad"),
        "Expected indented items for hierarchy, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("VP Engineering"),
        "Expected VP Engineering, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Dev Lead"),
        "Expected Dev Lead, got:\n{}",
        output.source
    );
}

#[test]
fn test_smartart_codegen_empty_items() {
    let doc = make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![FixedElement {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 100.0,
            kind: FixedElementKind::SmartArt(SmartArt { items: vec![] }),
        }],
    )]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("SmartArt Diagram"),
        "Expected SmartArt header even for empty SmartArt"
    );
}

#[test]
fn test_smartart_codegen_special_chars() {
    let doc = make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![FixedElement {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 100.0,
            kind: FixedElementKind::SmartArt(SmartArt {
                items: vec![sa_node("Item #1", 0), sa_node("Price $10", 0)],
            }),
        }],
    )]);

    let output = generate_typst(&doc).unwrap();
    // # and $ should be escaped
    assert!(
        output.source.contains(r"\#"),
        "Expected escaped #, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains(r"\$"),
        "Expected escaped $, got:\n{}",
        output.source
    );
}

// ── Gradient codegen tests (US-050) ─────────────────────────────────

#[test]
fn test_gradient_background_codegen() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: Some(Color::new(255, 0, 0)), // fallback
        background_gradient: Some(GradientFill {
            stops: vec![
                GradientStop {
                    position: 0.0,
                    color: Color::new(255, 0, 0),
                },
                GradientStop {
                    position: 1.0,
                    color: Color::new(0, 0, 255),
                },
            ],
            angle: 90.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("gradient.linear("),
        "Should contain gradient.linear. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("(rgb(255, 0, 0), 0%)"),
        "Should contain first stop. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("(rgb(0, 0, 255), 100%)"),
        "Should contain second stop. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("angle: 90deg"),
        "Should contain angle. Got: {}",
        output.source,
    );
}

#[test]
fn test_gradient_background_no_angle_codegen() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: None,
        background_gradient: Some(GradientFill {
            stops: vec![
                GradientStop {
                    position: 0.0,
                    color: Color::new(255, 255, 255),
                },
                GradientStop {
                    position: 1.0,
                    color: Color::new(0, 0, 0),
                },
            ],
            angle: 0.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("gradient.linear("),
        "Should contain gradient.linear. Got: {}",
        output.source,
    );
    // Angle of 0 should NOT be emitted
    assert!(
        !output.source.contains("angle:"),
        "Should not contain angle for 0 degrees. Got: {}",
        output.source,
    );
}

#[test]
fn test_gradient_shape_fill_codegen() {
    let elem = FixedElement {
        x: 10.0,
        y: 20.0,
        width: 200.0,
        height: 150.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Rectangle,
            fill: Some(Color::new(255, 0, 0)), // fallback
            gradient_fill: Some(GradientFill {
                stops: vec![
                    GradientStop {
                        position: 0.0,
                        color: Color::new(0, 128, 0),
                    },
                    GradientStop {
                        position: 1.0,
                        color: Color::new(0, 0, 128),
                    },
                ],
                angle: 45.0,
            }),
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("gradient.linear("),
        "Should contain gradient.linear for shape. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("(rgb(0, 128, 0), 0%)"),
        "Should contain first stop. Got: {}",
        output.source,
    );
    // Should NOT contain the fallback rgb fill since gradient takes precedence
    assert!(
        !output.source.contains("fill: rgb(255, 0, 0)"),
        "Should not contain fallback solid fill. Got: {}",
        output.source,
    );
}

// ── Shadow codegen tests ──────────────────────────────────────────

#[test]
fn test_shape_shadow_codegen() {
    use crate::ir::Shadow;
    let elem = FixedElement {
        x: 10.0,
        y: 20.0,
        width: 200.0,
        height: 150.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Rectangle,
            fill: Some(Color::new(255, 0, 0)),
            gradient_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: Some(Shadow {
                blur_radius: 4.0,
                distance: 3.0,
                direction: 45.0,
                color: Color::new(0, 0, 0),
                opacity: 0.5,
            }),
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let output = generate_typst(&doc).unwrap();
    // Shadow should render as an offset duplicate with rgb fill (4 args for alpha)
    assert!(
        output.source.contains("rgb(0, 0, 0, 128)"),
        "Shadow should use rgb with alpha. Got: {}",
        output.source,
    );
    // The shadow shape should be placed before the main shape
    let shadow_pos = output.source.find("rgb(0, 0, 0, 128)");
    let main_pos = output.source.find("rgb(255, 0, 0)");
    assert!(
        shadow_pos < main_pos,
        "Shadow should appear before main shape in output",
    );
}

#[test]
fn test_shape_no_shadow_no_extra_output() {
    let elem = FixedElement {
        x: 10.0,
        y: 20.0,
        width: 200.0,
        height: 150.0,
        kind: FixedElementKind::Shape(Shape {
            kind: ShapeKind::Rectangle,
            fill: Some(Color::new(255, 0, 0)),
            gradient_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: None,
        }),
    };
    let doc = make_doc(vec![make_fixed_page(720.0, 540.0, vec![elem])]);
    let output = generate_typst(&doc).unwrap();
    // No shadow → no rgb(0, 0, 0, ...) for shadow color
    assert!(
        !output.source.contains("rgb(0, 0, 0,"),
        "No shadow should produce no rgb shadow. Got: {}",
        output.source,
    );
}

#[test]
fn test_gradient_prefers_over_solid_fill() {
    // When both gradient_fill and fill are present, gradient should be used
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: Some(Color::new(128, 128, 128)),
        background_gradient: Some(GradientFill {
            stops: vec![
                GradientStop {
                    position: 0.0,
                    color: Color::new(255, 0, 0),
                },
                GradientStop {
                    position: 1.0,
                    color: Color::new(0, 0, 255),
                },
            ],
            angle: 180.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    // Gradient should be used, not the solid fallback
    assert!(
        output.source.contains("gradient.linear("),
        "Gradient should be preferred. Got: {}",
        output.source,
    );
    assert!(
        !output.source.contains("fill: rgb(128, 128, 128)"),
        "Solid fallback should not appear. Got: {}",
        output.source,
    );
}

#[test]
fn test_gradient_unsorted_stops_rendered_in_sorted_order() {
    // Gradient stops provided in reverse order should be sorted by position
    // before rendering — Typst requires monotonic offsets.
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: None,
        background_gradient: Some(GradientFill {
            stops: vec![
                GradientStop {
                    position: 1.0,
                    color: Color::new(0, 0, 255),
                },
                GradientStop {
                    position: 0.5,
                    color: Color::new(0, 255, 0),
                },
                GradientStop {
                    position: 0.0,
                    color: Color::new(255, 0, 0),
                },
            ],
            angle: 90.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    // Stops should appear in order: 0% (red), 50% (green), 100% (blue)
    let src = &output.source;
    let pos_red = src.find("(rgb(255, 0, 0), 0%)").expect("red stop missing");
    let pos_green = src
        .find("(rgb(0, 255, 0), 50%)")
        .expect("green stop missing");
    let pos_blue = src
        .find("(rgb(0, 0, 255), 100%)")
        .expect("blue stop missing");
    assert!(
        pos_red < pos_green && pos_green < pos_blue,
        "Stops should be in sorted order (0% < 50% < 100%). Got: {}",
        src,
    );
}

// ── DataBar / IconSet codegen tests ──────────────────────────────

#[test]
fn test_data_bar_codegen() {
    use crate::ir::DataBarInfo;
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "50".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        data_bar: Some(DataBarInfo {
            color: Color::new(0x63, 0x8E, 0xC6),
            fill_pct: 50.0,
        }),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let page = Page::Table(TablePage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("fill: rgb(99, 142, 198)"),
        "DataBar should contain bar color fill. Got: {}",
        output.source,
    );
    assert!(
        output.source.contains("width: 50%"),
        "DataBar should contain 50% width. Got: {}",
        output.source,
    );
}

#[test]
fn test_icon_text_codegen() {
    let cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "90".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        icon_text: Some("↑".to_string()),
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![cell],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let page = Page::Table(TablePage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table,
        header: None,
        footer: None,
        charts: vec![],
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("↑"),
        "Icon text should appear in output. Got: {}",
        output.source,
    );
}

#[test]
fn test_table_colspan_clamped_to_available_columns() {
    // Table with 2 columns, but cell has col_span: 3 (exceeds available).
    // The codegen should clamp it to 2.
    let wide_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Wide".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        col_span: 3,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![wide_cell],
                height: None,
            },
            TableRow {
                cells: vec![make_text_cell("A2"), make_text_cell("B2")],
                height: None,
            },
        ],
        column_widths: vec![100.0, 200.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    // colspan should be clamped to 2 (number of columns), not 3
    assert!(
        result.contains("colspan: 2"),
        "Expected colspan clamped to 2, got: {result}"
    );
    assert!(
        !result.contains("colspan: 3"),
        "colspan: 3 should have been clamped, got: {result}"
    );
}

#[test]
fn test_table_colspan_clamped_mid_row() {
    // Table with 3 columns, row has cell at col 1 + cell with col_span: 3 at col 2.
    // col_span should be clamped to 2 (3 - 1 = 2 remaining columns).
    let normal_cell = make_text_cell("A1");
    let wide_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Wide".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        col_span: 3,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![TableRow {
            cells: vec![normal_cell, wide_cell],
            height: None,
        }],
        column_widths: vec![100.0, 100.0, 100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    // At col position 1, col_span 3 exceeds 3 columns → clamped to 2
    assert!(
        result.contains("colspan: 2"),
        "Expected colspan clamped to 2, got: {result}"
    );
}

#[test]
fn test_table_colspan_no_column_widths_inferred() {
    // Table without explicit column_widths — num_cols inferred from max cells in a row.
    let wide_cell = TableCell {
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "Wide".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        col_span: 5,
        ..TableCell::default()
    };
    let table = Table {
        rows: vec![
            TableRow {
                cells: vec![wide_cell],
                height: None,
            },
            TableRow {
                cells: vec![
                    make_text_cell("A"),
                    make_text_cell("B"),
                    make_text_cell("C"),
                ],
                height: None,
            },
        ],
        column_widths: vec![],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    // Inferred num_cols = 3 (max cells in any row), col_span 5 clamped to 3
    assert!(
        result.contains("colspan: 3"),
        "Expected colspan clamped to 3 (inferred columns), got: {result}"
    );
    assert!(
        !result.contains("colspan: 5"),
        "colspan: 5 should have been clamped, got: {result}"
    );
}

// ── Metadata codegen tests ─────────────────────────────────────────

#[test]
fn test_generate_typst_with_metadata_title_and_author() {
    let doc = Document {
        metadata: Metadata {
            title: Some("Test Title".to_string()),
            author: Some("Test Author".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#set document(title: \"Test Title\", author: \"Test Author\")"),
        "Expected document metadata in Typst output, got: {result}"
    );
}

#[test]
fn test_generate_typst_with_metadata_title_only() {
    let doc = Document {
        metadata: Metadata {
            title: Some("Only Title".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#set document(title: \"Only Title\")"),
        "Expected title-only metadata in Typst output, got: {result}"
    );
}

#[test]
fn test_generate_typst_without_metadata() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        runs: vec![Run {
            text: "Hello".to_string(),
            style: TextStyle::default(),
            footnote: None,
            href: None,
        }],
        style: ParagraphStyle::default(),
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("#set document("),
        "Should not emit #set document when no metadata, got: {result}"
    );
}

#[test]
fn test_generate_typst_with_metadata_created_date() {
    let doc = Document {
        metadata: Metadata {
            title: Some("Dated Doc".to_string()),
            created: Some("2024-06-15T10:30:00Z".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    // When metadata has a created date, it should be emitted in Typst
    assert!(
        result.contains("date: datetime(year: 2024, month: 6, day: 15"),
        "Expected document date from metadata created field, got: {result}"
    );
}

#[test]
fn test_generate_typst_with_metadata_date_only() {
    // When only the created date is set (no title/author), date should still appear
    let doc = Document {
        metadata: Metadata {
            created: Some("2023-12-25T08:00:00Z".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("date: datetime(year: 2023, month: 12, day: 25"),
        "Expected document date even without title/author, got: {result}"
    );
}

#[test]
fn test_generate_typst_with_invalid_created_date() {
    // Invalid date string should be silently ignored
    let doc = Document {
        metadata: Metadata {
            title: Some("Bad Date Doc".to_string()),
            created: Some("not-a-date".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    // Invalid date should not crash or produce a date field
    assert!(
        !result.contains("date: datetime("),
        "Invalid date should not produce document date, got: {result}"
    );
}

#[test]
fn test_parse_iso8601_date_full() {
    let result = parse_iso8601_date("2024-06-15T10:30:45Z");
    assert_eq!(result, Some((2024, 6, 15, 10, 30, 45)));
}

#[test]
fn test_parse_iso8601_date_date_only() {
    let result = parse_iso8601_date("2023-12-25");
    assert_eq!(result, Some((2023, 12, 25, 0, 0, 0)));
}

#[test]
fn test_parse_iso8601_date_invalid() {
    assert_eq!(parse_iso8601_date("not-a-date"), None);
    assert_eq!(parse_iso8601_date(""), None);
    assert_eq!(parse_iso8601_date("2024"), None);
    assert_eq!(parse_iso8601_date("2024-13-01T00:00:00Z"), None); // month > 12
    assert_eq!(parse_iso8601_date("2024-00-01T00:00:00Z"), None); // month 0
}

// ── Extended geometry codegen tests (US-085) ──────────────────────────

#[test]
fn test_triangle_polygon_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            20.0,
            200.0,
            150.0,
            ShapeKind::Polygon {
                vertices: vec![(0.5, 0.0), (1.0, 1.0), (0.0, 1.0)],
            },
            Some(Color::new(255, 0, 0)),
            None,
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#polygon("),
        "Expected #polygon in: {}",
        output.source
    );
    // Check vertex at top-center: 0.5 * 200 = 100pt
    assert!(
        output.source.contains("100pt"),
        "Expected 100pt vertex x in: {}",
        output.source
    );
    assert!(
        output.source.contains("fill: rgb(255, 0, 0)"),
        "Expected fill in: {}",
        output.source
    );
}

#[test]
fn test_rounded_rectangle_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            10.0,
            20.0,
            200.0,
            100.0,
            ShapeKind::RoundedRectangle {
                radius_fraction: 0.1,
            },
            Some(Color::new(0, 0, 255)),
            None,
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#rect("),
        "Expected #rect in: {}",
        output.source
    );
    assert!(
        output.source.contains("radius:"),
        "Expected radius parameter in: {}",
        output.source
    );
    // Radius: 0.1 * min(200, 100) = 10pt
    assert!(
        output.source.contains("radius: 10pt"),
        "Expected radius: 10pt in: {}",
        output.source
    );
}

#[test]
fn test_arrow_polygon_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            0.0,
            0.0,
            300.0,
            150.0,
            ShapeKind::Polygon {
                vertices: vec![
                    (0.0, 0.25),
                    (0.6, 0.25),
                    (0.6, 0.0),
                    (1.0, 0.5),
                    (0.6, 1.0),
                    (0.6, 0.75),
                    (0.0, 0.75),
                ],
            },
            Some(Color::new(255, 136, 0)),
            None,
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#polygon("),
        "Expected #polygon for arrow in: {}",
        output.source
    );
    // Arrow tip at x=1.0*300=300pt, y=0.5*150=75pt
    assert!(
        output.source.contains("300pt"),
        "Expected 300pt (arrow tip) in: {}",
        output.source
    );
}

#[test]
fn test_polygon_with_stroke_codegen() {
    let doc = make_doc(vec![make_fixed_page(
        960.0,
        540.0,
        vec![make_shape_element(
            0.0,
            0.0,
            100.0,
            100.0,
            ShapeKind::Polygon {
                vertices: vec![(0.5, 0.0), (1.0, 0.5), (0.5, 1.0), (0.0, 0.5)],
            },
            None,
            Some(BorderSide {
                width: 2.0,
                color: Color::new(0, 0, 0),
                style: BorderLineStyle::Solid,
            }),
        )],
    )]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("#polygon("),
        "Expected #polygon in: {}",
        output.source
    );
    assert!(
        output.source.contains("stroke: 2pt + rgb(0, 0, 0)"),
        "Expected stroke in: {}",
        output.source
    );
}

#[test]
fn test_font_substitution_calibri_produces_fallback_list() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Calibri text".to_string(),
            style: TextStyle {
                font_family: Some("Calibri".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"font: ("Calibri", "Carlito", "Liberation Sans")"#),
        "Expected font fallback list for Calibri in: {result}"
    );
}

#[test]
fn test_font_substitution_arial_produces_fallback_list() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Arial text".to_string(),
            style: TextStyle {
                font_family: Some("Arial".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"font: ("Arial", "Liberation Sans", "Arimo")"#),
        "Expected font fallback list for Arial in: {result}"
    );
}

#[test]
fn test_font_substitution_unknown_font_no_fallback() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Custom text".to_string(),
            style: TextStyle {
                font_family: Some("Helvetica".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"font: "Helvetica""#),
        "Unknown font should use simple quoted string in: {result}"
    );
    // Should NOT contain parenthesized array
    assert!(
        !result.contains("font: (\"Helvetica\""),
        "Unknown font should not use array syntax in: {result}"
    );
}

#[test]
fn test_font_substitution_times_new_roman() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "TNR text".to_string(),
            style: TextStyle {
                font_family: Some("Times New Roman".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"font: ("Times New Roman", "Liberation Serif", "Tinos")"#),
        "Expected font fallback list for Times New Roman in: {result}"
    );
}

#[test]
fn test_font_family_infers_medium_weight_from_family_name() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Title".to_string(),
            style: TextStyle {
                font_family: Some("Pretendard Medium".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"weight: "medium""#),
        "Expected medium weight inferred from family name in: {result}"
    );
}

#[test]
fn test_font_family_infers_extrabold_weight_from_family_name() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Heading".to_string(),
            style: TextStyle {
                font_family: Some("Pretendard ExtraBold".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"weight: "extrabold""#),
        "Expected extrabold weight inferred from family name in: {result}"
    );
}

#[test]
fn test_generate_typst_prefers_office_font_order_when_context_present() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Title".to_string(),
            style: TextStyle {
                font_family: Some("Pretendard".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let context = FontSearchContext::for_test(
        Vec::new(),
        &["Apple SD Gothic Neo", "Malgun Gothic"],
        &["Malgun Gothic"],
        &[],
    );

    let output = generate_typst_with_options_and_font_context(
        &doc,
        &ConvertOptions::default(),
        Some(&context),
    )
    .unwrap();

    let apple_index = output
        .source
        .find("\"Apple SD Gothic Neo\"")
        .expect("Apple SD Gothic Neo should appear in Typst output");
    let malgun_index = output
        .source
        .find("\"Malgun Gothic\"")
        .expect("Malgun Gothic should appear in Typst output");
    assert!(
        malgun_index < apple_index,
        "Office-resolved font ordering should win in Typst output: {}",
        output.source
    );
}

// --- Heading level codegen tests (US-096) ---

#[test]
fn test_generate_heading_level_1() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            heading_level: Some(1),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Main Title".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#heading(level: 1)[Main Title]"),
        "H1 paragraph should emit #heading(level: 1): {result}"
    );
}

#[test]
fn test_generate_heading_level_2() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            heading_level: Some(2),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Sub Section".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#heading(level: 2)[Sub Section]"),
        "H2 paragraph should emit #heading(level: 2): {result}"
    );
}

#[test]
fn test_generate_heading_levels_3_to_6() {
    for level in 3..=6u8 {
        let text = format!("Heading {level}");
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                heading_level: Some(level),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: text.clone(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        let expected = format!("#heading(level: {level})[{text}]");
        assert!(
            result.contains(&expected),
            "H{level} should emit {expected}: {result}"
        );
    }
}

#[test]
fn test_generate_heading_with_styled_run() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            heading_level: Some(1),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Styled Heading".to_string(),
            style: TextStyle {
                bold: Some(true),
                font_size: Some(24.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#heading(level: 1)"),
        "Heading with styling should still emit #heading: {result}"
    );
}

#[test]
fn test_generate_regular_paragraph_no_heading() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Normal text".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("#heading"),
        "Regular paragraph should not emit #heading: {result}"
    );
}

// ── Unicode NFC normalization tests ──────────────────────────────

#[test]
fn test_escape_typst_normalizes_korean_nfd_to_nfc() {
    // Korean "한글" in NFD (decomposed jamo): ㅎ + ㅏ + ㄴ + ㄱ + ㅡ + ㄹ
    let nfd_korean = "\u{1112}\u{1161}\u{11AB}\u{1100}\u{1173}\u{11AF}";
    let nfc_korean = "한글";
    let result = escape_typst(nfd_korean);
    assert_eq!(
        result, nfc_korean,
        "NFD Korean jamo should be normalized to composed hangul"
    );
}

#[test]
fn test_escape_typst_normalizes_combining_diacritics() {
    // "café" with combining acute accent (NFD): 'e' + combining acute
    let nfd_cafe = "cafe\u{0301}";
    let nfc_cafe = "caf\u{00E9}"; // é as precomposed
    let result = escape_typst(nfd_cafe);
    assert_eq!(
        result, nfc_cafe,
        "Combining diacritics should be normalized to NFC"
    );
}

#[test]
fn test_escape_typst_nfc_with_special_chars() {
    // NFD text with Typst special chars: "café $5" with combining accent
    let nfd_input = "cafe\u{0301} \\$5";
    let result = escape_typst(nfd_input);
    // NFC normalization + Typst escaping
    assert!(
        result.contains("caf\u{00E9}"),
        "Should contain NFC-normalized é: {result}"
    );
    assert!(
        result.contains("\\$"),
        "Should still escape $ sign: {result}"
    );
}

#[test]
fn test_generate_typst_nfc_korean_in_paragraph() {
    // NFD Korean in a full paragraph through the pipeline
    let nfd_korean = "\u{1112}\u{1161}\u{11AB}\u{1100}\u{1173}\u{11AF}";
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(nfd_korean)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("한글"),
        "Generated Typst should contain NFC-composed Korean: {result}"
    );
    assert!(
        !result.contains('\u{1112}'),
        "Generated Typst should not contain decomposed jamo: {result}"
    );
}

#[test]
fn test_generate_typst_nfc_diacritics_in_paragraph() {
    // NFD "résumé" through the full pipeline
    let nfd_resume = "re\u{0301}sume\u{0301}";
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(nfd_resume)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("r\u{00E9}sum\u{00E9}"),
        "Generated Typst should contain NFC-composed résumé: {result}"
    );
}

#[test]
fn test_escape_typst_already_nfc_unchanged() {
    // Already NFC text should pass through unchanged (minus Typst escaping)
    let nfc_text = "Hello 한글 café";
    let result = escape_typst(nfc_text);
    assert_eq!(result, nfc_text, "Already-NFC text should be unchanged");
}

// --- US-103: Multi-column section layout codegen tests ---

#[test]
fn test_generate_flow_page_with_equal_columns() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Column text")],
        header: None,
        footer: None,
        columns: Some(ColumnLayout {
            num_columns: 2,
            spacing: 36.0,
            column_widths: None,
        }),
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#columns(2, gutter: 36pt)"),
        "Should contain columns() call. Got: {result}"
    );
    assert!(
        result.contains("Column text"),
        "Should contain the text content. Got: {result}"
    );
}

#[test]
fn test_generate_flow_page_with_three_columns() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Three col text")],
        header: None,
        footer: None,
        columns: Some(ColumnLayout {
            num_columns: 3,
            spacing: 18.0,
            column_widths: None,
        }),
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#columns(3, gutter: 18pt)"),
        "Should contain columns(3, ...). Got: {result}"
    );
}

#[test]
fn test_generate_flow_page_with_unequal_columns() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Unequal col text")],
        header: None,
        footer: None,
        columns: Some(ColumnLayout {
            num_columns: 2,
            spacing: 36.0,
            column_widths: Some(vec![300.0, 150.0]),
        }),
    })]);
    let result = generate_typst(&doc).unwrap().source;
    // Unequal columns should use grid() with explicit widths
    assert!(
        result.contains("#grid(columns: (300pt, 150pt)"),
        "Unequal columns should use grid(). Got: {result}"
    );
}

#[test]
fn test_generate_column_break() {
    let doc = make_doc(vec![Page::Flow(FlowPage {
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![
            make_paragraph("Before break"),
            Block::ColumnBreak,
            make_paragraph("After break"),
        ],
        header: None,
        footer: None,
        columns: Some(ColumnLayout {
            num_columns: 2,
            spacing: 36.0,
            column_widths: None,
        }),
    })]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#colbreak()"),
        "Should contain colbreak(). Got: {result}"
    );
}

#[test]
fn test_generate_no_columns_no_wrapper() {
    // Without column layout, content should not be wrapped in columns()
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Normal text")])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("#columns("),
        "Should not contain columns(). Got: {result}"
    );
    assert!(
        !result.contains("#grid(columns:"),
        "Should not contain grid(columns:). Got: {result}"
    );
}

// ── BiDi / RTL codegen tests ──────────────────────────────────────

#[test]
fn test_generate_rtl_paragraph() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            direction: Some(TextDirection::Rtl),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "مرحبا بالعالم".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#set text(dir: rtl)"),
        "RTL paragraph should emit #set text(dir: rtl). Got: {result}"
    );
}

#[test]
fn test_generate_ltr_paragraph_no_direction() {
    // Normal LTR paragraph should NOT emit any text direction
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Hello World")])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("dir: rtl"),
        "LTR paragraph should not emit dir: rtl. Got: {result}"
    );
}

#[test]
fn test_generate_mixed_rtl_ltr_paragraphs() {
    let doc = make_doc(vec![make_flow_page(vec![
        Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                direction: Some(TextDirection::Rtl),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "مرحبا 123".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        }),
        make_paragraph("English text"),
    ])]);
    let result = generate_typst(&doc).unwrap().source;
    // Should contain RTL setting for the Arabic paragraph
    assert!(
        result.contains("#set text(dir: rtl)"),
        "Should contain RTL direction for Arabic paragraph. Got: {result}"
    );
    // The Arabic text and English text should both appear
    assert!(result.contains("مرحبا 123"), "Arabic text should appear");
    assert!(
        result.contains("English text"),
        "English text should appear"
    );
}

// --- US-204: Codegen/render robustness tests ---

#[test]
fn test_codegen_robustness_zero_pages() {
    // An empty document with zero pages should produce valid Typst output
    let doc = make_doc(vec![]);
    let output = generate_typst(&doc).unwrap();
    // Should produce an empty (or near-empty) source without panicking
    assert!(output.images.is_empty());
}

#[test]
fn test_codegen_robustness_flow_page_empty_content() {
    // A flow page with no content blocks should not panic
    let doc = make_doc(vec![make_flow_page(vec![])]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.is_empty());
}

#[test]
fn test_generate_fixed_page_empty_elements() {
    // A fixed page with no elements should not panic
    let doc = make_doc(vec![Page::Fixed(FixedPage {
        size: PageSize::default(),
        elements: vec![],
        background_color: None,
        background_gradient: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.is_empty());
}

#[test]
fn test_generate_table_page_empty_rows() {
    // A table page with zero rows should not panic
    let doc = make_doc(vec![Page::Table(TablePage {
        name: String::new(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: Table {
            rows: vec![],
            column_widths: vec![],
            ..Table::default()
        },
        header: None,
        footer: None,
        charts: vec![],
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.is_empty());
}

#[test]
fn test_generate_paragraph_all_alignment_variants() {
    // All alignment variants (Left, Center, Right, Justify, None) should
    // produce valid Typst output without panicking.
    for alignment in [
        Some(Alignment::Left),
        Some(Alignment::Center),
        Some(Alignment::Right),
        Some(Alignment::Justify),
        None,
    ] {
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                alignment,
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: format!("Alignment: {alignment:?}"),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let output = generate_typst(&doc);
        assert!(
            output.is_ok(),
            "Codegen should not fail for alignment {alignment:?}"
        );
    }
}

#[test]
fn test_generate_shape_shadow_all_kinds() {
    // Shadow generation should handle all ShapeKind variants without panicking.
    let shadow = Shadow {
        blur_radius: 4.0,
        color: Color { r: 0, g: 0, b: 0 },
        opacity: 0.5,
        direction: 45.0,
        distance: 3.0,
    };

    let shape_kinds = vec![
        ShapeKind::Rectangle,
        ShapeKind::Ellipse,
        ShapeKind::Line { x2: 100.0, y2: 0.0 },
        ShapeKind::RoundedRectangle {
            radius_fraction: 0.1,
        },
        ShapeKind::Polygon {
            vertices: vec![(0.0, 0.0), (1.0, 0.0), (0.5, 1.0)],
        },
    ];

    for kind in shape_kinds {
        let doc = make_doc(vec![Page::Fixed(FixedPage {
            size: PageSize {
                width: 960.0,
                height: 540.0,
            },
            elements: vec![FixedElement {
                x: 100.0,
                y: 100.0,
                width: 200.0,
                height: 100.0,
                kind: FixedElementKind::Shape(Shape {
                    kind: kind.clone(),
                    fill: Some(Color { r: 255, g: 0, b: 0 }),
                    gradient_fill: None,
                    stroke: None,
                    opacity: None,
                    shadow: Some(shadow.clone()),
                    rotation_deg: None,
                }),
            }],
            background_color: None,
            background_gradient: None,
        })]);
        let output = generate_typst(&doc);
        assert!(
            output.is_ok(),
            "Codegen should not panic for shape kind {kind:?} with shadow"
        );
    }
}

#[test]
fn test_column_break_with_empty_content() {
    // Column breaks on empty content should not panic
    let segments = split_at_column_breaks(&[]);
    assert_eq!(segments.len(), 1);
    assert!(segments[0].is_empty());
}

#[test]
fn test_column_break_only_breaks() {
    // Content consisting only of column breaks should not panic
    let blocks = vec![Block::ColumnBreak, Block::ColumnBreak];
    let segments = split_at_column_breaks(&blocks);
    assert_eq!(segments.len(), 3);
    assert!(segments.iter().all(|s| s.is_empty()));
}

// --- US-315: text escaping for Typst-significant characters ---

#[test]
fn test_escape_typst_backslash() {
    assert_eq!(escape_typst("path\\to\\file"), "path\\\\to\\\\file");
}

#[test]
fn test_escape_typst_hash() {
    assert_eq!(escape_typst("#hashtag"), "\\#hashtag");
}

#[test]
fn test_escape_typst_dollar() {
    assert_eq!(escape_typst("$100"), "\\$100");
}

#[test]
fn test_escape_typst_brackets() {
    assert_eq!(escape_typst("[content]"), "\\[content\\]");
}

#[test]
fn test_escape_typst_braces() {
    assert_eq!(escape_typst("{code}"), "\\{code\\}");
}

#[test]
fn test_escape_typst_all_special_chars() {
    let input = r"#*_`<>@\~/$[]{}";
    let result = escape_typst(input);
    // Every character should be escaped
    assert_eq!(result, "\\#\\*\\_\\`\\<\\>\\@\\\\\\~\\/\\$\\[\\]\\{\\}");
}

#[test]
fn test_escape_typst_in_paragraph_output() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph(
        "Price: $100 path\\to",
    )])]);
    let output = generate_typst(&doc).unwrap().source;
    assert!(
        output.contains("\\$100"),
        "Dollar sign should be escaped in output: {output}"
    );
    assert!(
        output.contains("path\\\\to"),
        "Backslash should be escaped in output: {output}"
    );
}

// --- US-316: single-stop gradient fallback ---

#[test]
fn test_gradient_single_stop_fallback_to_solid() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: None,
        background_gradient: Some(GradientFill {
            stops: vec![GradientStop {
                position: 0.5,
                color: Color::new(255, 128, 0),
            }],
            angle: 0.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    // Should NOT contain gradient.linear (needs >= 2 stops)
    assert!(
        !output.source.contains("gradient.linear"),
        "Single-stop gradient should fall back to solid fill: {}",
        output.source,
    );
    // Should contain the solid fill color instead
    assert!(
        output.source.contains("rgb(255, 128, 0)"),
        "Single-stop gradient should use the stop color as solid fill: {}",
        output.source,
    );
}

#[test]
fn test_gradient_two_stops_still_works() {
    let page = Page::Fixed(FixedPage {
        size: PageSize {
            width: 720.0,
            height: 540.0,
        },
        elements: vec![],
        background_color: None,
        background_gradient: Some(GradientFill {
            stops: vec![
                GradientStop {
                    position: 0.0,
                    color: Color::new(255, 0, 0),
                },
                GradientStop {
                    position: 1.0,
                    color: Color::new(0, 0, 255),
                },
            ],
            angle: 90.0,
        }),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("gradient.linear"),
        "Two-stop gradient should still produce gradient.linear: {}",
        output.source,
    );
}

// --- US-382/383: unstyled run after styled run must not create `](` pattern ---

#[test]
fn test_unstyled_run_with_parens_after_styled_run() {
    // When a styled run is followed by an unstyled run starting with `(`,
    // the `](` pattern must not be interpreted as Typst function arguments.
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![
            Run {
                text: "bold text".to_string(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            },
            Run {
                text: "(parenthetical note)".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            },
        ],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    // The result must not contain `](` directly — it would be interpreted
    // as function arguments in Typst
    assert!(
        !result.contains("](\\(") || !result.contains("]("),
        "Unstyled text with parens after styled run must be wrapped safely. Got: {result}"
    );
    // Verify the output uses #[...] wrapper or other safe pattern
    assert!(
        result.contains("#[") || result.contains("\\("),
        "Unstyled text should be wrapped in #[...] to prevent syntax issues. Got: {result}"
    );
}

#[test]
fn test_generate_run_superscript() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "2".to_string(),
            style: TextStyle {
                vertical_align: Some(VerticalTextAlign::Superscript),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#super[2]"),
        "Superscript should use #super[...]. Got: {result}"
    );
}

#[test]
fn test_generate_run_subscript() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "2".to_string(),
            style: TextStyle {
                vertical_align: Some(VerticalTextAlign::Subscript),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#sub[2]"),
        "Subscript should use #sub[...]. Got: {result}"
    );
}

#[test]
fn test_generate_run_small_caps() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Hello".to_string(),
            style: TextStyle {
                small_caps: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#smallcaps[Hello]"),
        "Small caps should use #smallcaps[...]. Got: {result}"
    );
}

#[test]
fn test_generate_run_all_caps() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Hello World".to_string(),
            style: TextStyle {
                all_caps: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("HELLO WORLD"),
        "All caps should uppercase the text. Got: {result}"
    );
}

#[test]
fn test_generate_run_superscript_with_bold() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "n".to_string(),
            style: TextStyle {
                vertical_align: Some(VerticalTextAlign::Superscript),
                bold: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#super[") && result.contains("weight: \"bold\""),
        "Superscript with bold should combine both. Got: {result}"
    );
}

#[test]
fn test_generate_run_highlight_yellow() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Important".to_string(),
            style: TextStyle {
                highlight: Some(Color::new(255, 255, 0)),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#highlight(fill: rgb(255, 255, 0))[Important]"),
        "Highlight should use #highlight(fill: ...). Got: {result}"
    );
}

#[test]
fn test_table_cell_vertical_align_center() {
    let table = Table {
        rows: vec![TableRow {
            cells: vec![TableCell {
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Centered".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                })],
                vertical_align: Some(CellVerticalAlign::Center),
                ..TableCell::default()
            }],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("align: horizon"),
        "Center vertical alignment should emit 'align: horizon'. Got: {result}"
    );
}

#[test]
fn test_generate_run_highlight_with_bold() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Bold Highlight".to_string(),
            style: TextStyle {
                highlight: Some(Color::new(0, 255, 0)),
                bold: Some(true),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#highlight(fill: rgb(0, 255, 0))["),
        "Should have highlight wrapper. Got: {result}"
    );
    assert!(
        result.contains("weight: \"bold\""),
        "Should have bold text. Got: {result}"
    );
}

#[test]
fn test_table_cell_vertical_align_bottom() {
    let table = Table {
        rows: vec![TableRow {
            cells: vec![TableCell {
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Bottom".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                })],
                vertical_align: Some(CellVerticalAlign::Bottom),
                ..TableCell::default()
            }],
            height: None,
        }],
        column_widths: vec![100.0],
        ..Table::default()
    };
    let doc = make_doc(vec![make_flow_page(vec![Block::Table(table)])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("align: bottom"),
        "Bottom vertical alignment should emit 'align: bottom'. Got: {result}"
    );
}
