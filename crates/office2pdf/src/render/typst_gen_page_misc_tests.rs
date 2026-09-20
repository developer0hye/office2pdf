use super::*;
use crate::ir::ChartAreaOutline;

#[test]
fn test_generate_flow_page_with_text_header() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body text")],
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Document Title".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("header:"));
    assert!(output.source.contains("Document Title"));
}

#[test]
fn test_generate_flow_page_with_page_number_footer() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body text")],
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: Some(35.4),
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![
                    HFInline::Run(Run {
                        text: "Page ".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }),
                    HFInline::PageNumber(TextStyle::default()),
                ],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("footer:"));
    assert!(output.source.contains(r#"counter(page).display("1")"#));
    assert!(output.source.contains("Page "));
    // Word pins the footer's bottom `w:footer` points above the page edge.
    assert!(output.source.contains("footer-descent: 0pt"));
    assert!(output.source.contains("block(width: 100%, height: 36.6pt)"));
    assert!(output.source.contains("place(bottom"));
}

#[test]
fn test_generate_footer_with_compound_border_and_right_positioned_tab() {
    use crate::ir::{
        BorderSide, CellBorder, HFInline, HeaderFooter, HeaderFooterParagraph, PositionedTab,
        PositionedTabAlignment, PositionedTabRelativeTo,
    };

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![
                    HFInline::Run(Run {
                        text: "Left".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }),
                    HFInline::PositionedTab(PositionedTab {
                        alignment: PositionedTabAlignment::Right,
                        relative_to: PositionedTabRelativeTo::Margin,
                        leader: TabLeader::None,
                    }),
                    HFInline::Run(Run {
                        text: "Page ".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }),
                    HFInline::PageNumber(TextStyle::default()),
                ],
                border: Some(CellBorder {
                    top: Some(BorderSide {
                        width: 3.0,
                        color: Color::new(0x62, 0x24, 0x23),
                        style: BorderLineStyle::Double,
                        join: LineJoin::Round,
                    }),
                    bottom: None,
                    left: None,
                    right: None,
                }),
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("#grid(columns: (1fr, auto)"));
    assert!(output.source.contains("rgb(98, 36, 35)"));
    assert_eq!(output.source.matches("line(length: 100%").count(), 2);
}

#[test]
fn a_page_anchored_footer_frame_paints_below_body_content() {
    use crate::ir::{
        FrameAnchor, HFInline, HeaderFooter, HeaderFooterFrame, HeaderFooterParagraph,
    };

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Framed footer".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: Some(HeaderFooterFrame {
                    wraps_text: true,
                    x: Some(71.8),
                    y: Some(198.5),
                    width: None,
                    height: None,
                    horizontal_anchor: FrameAnchor::Page,
                    vertical_anchor: FrameAnchor::Page,
                    horizontal_align: None,
                    vertical_align: None,
                    inset_left: 0.0,
                    inset_top: 0.0,
                    bottom_offset: None,
                }),
            }],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("background: ["));
    assert!(
        output
            .source
            .contains("#place(top + left, dx: 71.8pt, dy: 198.5pt)")
    );
    assert!(!output.source.contains("foreground: ["));
    assert!(!output.source.contains("footer:"));
}

/// A page number inside a page-anchored frame has to compile.
///
/// The frame is emitted as a `#place` at document level, where the page
/// counter has no context of its own, so a bare `#counter(page).display()`
/// makes Typst abort the whole conversion (issue #788).
#[test]
fn test_page_anchored_frame_page_number_compiles() {
    use crate::ir::{
        FrameAnchor, HFInline, HeaderFooter, HeaderFooterFrame, HeaderFooterParagraph,
    };

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::PageNumber(TextStyle::default())],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: Some(HeaderFooterFrame {
                    wraps_text: true,
                    x: Some(50.0),
                    y: Some(25.0),
                    width: None,
                    height: None,
                    horizontal_anchor: FrameAnchor::Page,
                    vertical_anchor: FrameAnchor::Page,
                    horizontal_align: None,
                    vertical_align: None,
                    inset_left: 0.0,
                    inset_top: 0.0,
                    bottom_offset: None,
                }),
            }],
        }),
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("#place(top + left, dx: 50pt, dy: 25pt)"),
        "the frame is still placed at document level: {}",
        output.source
    );
    assert!(
        output.source.contains("#context counter(page)"),
        "the counter needs its own context there: {}",
        output.source
    );
    crate::render::pdf::compile_to_pdf(&output.source, &output.images, None, &[], false, false)
        .expect("a framed page number must compile");
}

#[test]
fn test_generate_flow_page_with_header_and_footer() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Header".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::PageNumber(TextStyle::default())],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("header:") && output.source.contains("footer:"));
}

#[test]
fn test_generate_flow_page_without_header_footer() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Body")])]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.contains("header:"));
    assert!(!output.source.contains("footer:"));
}

#[test]
fn test_generate_typst_inserts_pagebreak_between_flow_pages() {
    let first = Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("First section")],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    });
    let second = Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Second section")],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    });

    let output = generate_typst(&make_doc(vec![first, second])).unwrap();
    let pagebreak_count = output.source.matches("#pagebreak()").count();

    assert_eq!(pagebreak_count, 1);
}

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
    assert!(output.source.contains("fill: rgb(255, 0, 0)"));
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
        output.source.contains("fill: white"),
        "Expected fill: white for no-background slide, got:\n{}",
        output.source
    );
}

#[test]
fn test_fixed_page_table_element() {
    let table = Table {
        rows: vec![TableRow {
            minimum_height: None,
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
            .contains(r#"#link("https://example.com")[Click me]"#)
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
    assert!(output.source.contains(r#"#link("https://example.com")["#));
    assert!(output.source.contains("#text(weight: \"bold\")"));
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
    assert!(output.source.contains("Visit "));
    assert!(
        output
            .source
            .contains(r#"#link("https://rust-lang.org")[Rust]"#)
    );
    assert!(output.source.contains(" for more."));
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
            .contains(r#"#link("https://example.com/path?q=1&r=2")[Link]"#)
    );
}

/// A note's content run: a note carries styled runs, not one string.
fn note_run(text: &str) -> Run {
    Run {
        text: text.to_string(),
        style: TextStyle::default(),
        href: None,
        footnote: None,
    }
}

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
                footnote: Some(vec![note_run("This is a footnote.")]),
            },
        ],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("#footnote[This is a footnote.]"));
}

#[test]
fn test_footnote_with_special_chars() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: String::new(),
            style: TextStyle::default(),
            href: None,
            footnote: Some(vec![note_run("Note with #special *chars*")]),
        }],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains(r"#footnote[Note with \#special \*chars\*]")
    );
}

#[test]
fn test_table_page_with_header() {
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: make_simple_table(vec![vec!["A"]]),
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
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
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("header: ["));
    assert!(output.source.contains("My Header"));
}

#[test]
fn test_table_page_with_page_number_footer() {
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: make_simple_table(vec![vec!["A"]]),
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
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
                    HFInline::PageNumber(TextStyle::default()),
                    HFInline::Run(Run {
                        text: " of ".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }),
                    HFInline::TotalPages(TextStyle::default()),
                ],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("footer: context ["));
    assert!(
        output
            .source
            .contains(r#"#context counter(page).display("1")"#)
    );
    assert!(
        output
            .source
            .contains("#context counter(page).final().first()")
    );
}

/// A sheet page carrying a footer seat, for the seating tests below.
///
/// `footer_margin_pt` is `<pageMargins>/@footer` in paper points, as the
/// parser carries it; the footer is one plain 8pt run.
#[cfg(not(target_arch = "wasm32"))]
fn sheet_page_with_seated_footer(footer_margin_pt: Option<f64>, family: Option<&str>) -> Page {
    sheet_page_with_footer_sections(
        PageSize::default(),
        54.0,
        footer_margin_pt,
        None,
        vec![(
            Alignment::Left,
            false,
            vec![hf_run("Sensitivity: Internal", family, 8.0)],
        )],
    )
}

/// One footer run of `family` at `size_pt`.
#[cfg(not(target_arch = "wasm32"))]
fn hf_run(text: &str, family: Option<&str>, size_pt: f64) -> Run {
    Run {
        text: text.to_string(),
        style: TextStyle {
            font_family: family.map(str::to_string),
            font_size: Some(size_pt),
            ..TextStyle::default()
        },
        href: None,
        footnote: None,
    }
}

/// A sheet page whose footer has one paragraph per `(alignment, is_rich,
/// runs)` section, seated on `footer_margin_pt` with the given fit scale.
#[cfg(not(target_arch = "wasm32"))]
fn sheet_page_with_footer_sections(
    size: PageSize,
    bottom_margin_pt: f64,
    footer_margin_pt: Option<f64>,
    sheet_print_scale: Option<f64>,
    sections: Vec<(Alignment, bool, Vec<Run>)>,
) -> Page {
    Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size,
        margins: Margins {
            top: 54.0,
            bottom: bottom_margin_pt,
            left: 50.0,
            right: 50.0,
        },
        table: make_simple_table(vec![vec!["A"]]),
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: footer_margin_pt,
            sheet_print_scale,
            paragraphs: sections
                .into_iter()
                .map(
                    |(alignment, sheet_section_is_rich, runs)| HeaderFooterParagraph {
                        style: ParagraphStyle {
                            alignment: Some(alignment),
                            ..ParagraphStyle::default()
                        },
                        elements: runs.into_iter().map(HFInline::Run).collect(),
                        border: None,
                        border_space: None,
                        frame: None,
                        sheet_section_is_rich,
                    },
                )
                .collect(),
        }),
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    })
}

/// Each worksheet section can use the page width, even with all three present.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn sheet_header_footer_sections_preserve_page_width_and_explicit_breaks() {
    for is_header in [false, true] {
        for has_break in [false, true] {
            let text = if has_break {
                "Generated by Example Reporting System\nSecond line"
            } else {
                "Generated by Example Reporting System"
            };
            for alignment in [Alignment::Left, Alignment::Center, Alignment::Right] {
                let sections = [Alignment::Left, Alignment::Center, Alignment::Right]
                    .into_iter()
                    .map(|slot| {
                        let text = if slot == alignment { text } else { "Label" };
                        (slot, false, vec![hf_run(text, Some("Arial"), 11.0)])
                    })
                    .collect();
                let mut page = sheet_page_with_footer_sections(
                    PageSize {
                        width: 612.0,
                        height: 792.0,
                    },
                    54.0,
                    Some(21.6),
                    None,
                    sections,
                );
                if is_header {
                    let Page::Sheet(sheet) = &mut page else {
                        unreachable!()
                    };
                    sheet.header = sheet.footer.take();
                }
                let source = generate_typst(&make_doc(vec![page])).unwrap().source;
                let mut single_page = sheet_page_with_footer_sections(
                    PageSize {
                        width: 612.0,
                        height: 792.0,
                    },
                    54.0,
                    Some(21.6),
                    None,
                    vec![(alignment, false, vec![hf_run(text, Some("Arial"), 11.0)])],
                );
                if is_header {
                    let Page::Sheet(sheet) = &mut single_page else {
                        unreachable!()
                    };
                    sheet.header = sheet.footer.take();
                }
                let single_source = generate_typst(&make_doc(vec![single_page])).unwrap().source;
                let single_runs =
                    crate::render::pdf::compiled_text_runs(&single_source, 0).unwrap();
                let all_runs = crate::render::pdf::compiled_text_runs(&source, 0).unwrap();
                let single = single_runs
                    .iter()
                    .find(|run| run.text.contains("Generated"))
                    .unwrap();
                let shared = all_runs
                    .iter()
                    .find(|run| run.text.contains("Generated"))
                    .unwrap();
                assert!(
                    (single.left_pt - shared.left_pt).abs() < 0.01,
                    "{alignment:?} section horizontal anchor changed: {single:?} vs {shared:?}"
                );
                assert!(
                    (single.baseline_pt - shared.baseline_pt).abs() < 0.01,
                    "{alignment:?} section baseline changed: {single:?} vs {shared:?}"
                );
                if has_break {
                    let second = compiled_baseline_of(&source, "Second");
                    let single_second = compiled_baseline_of(&single_source, "Second");
                    assert!((second - single_second).abs() < 0.01);
                    assert!(
                        second > shared.baseline_pt + 5.0,
                        "explicit break must remain"
                    );
                }
                let first = compiled_baseline_of(&source, "Generated");
                let last = compiled_baseline_of(&source, "System");
                assert!(
                    (first - last).abs() < 0.01,
                    "{alignment:?} section must stay on one line: {first} vs {last}"
                );
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn multiline_sheet_header_uses_native_line_advance() {
    for (family, size, advance) in [
        ("Arial", 8.0, 11.0),
        ("Arial", 11.0, 14.0),
        ("Arial", 12.0, 15.0),
        ("Arial", 14.0, 17.0),
        ("Arial", 24.0, 29.0),
        ("Calibri", 11.0, 14.0),
        ("Aptos", 11.0, 14.0),
        ("Times New Roman", 11.0, 14.0),
        ("Malgun Gothic", 11.0, 17.0),
    ] {
        for line_count in [2, 3] {
            let mut page = sheet_page_with_footer_sections(
                PageSize {
                    width: 612.0,
                    height: 792.0,
                },
                72.0,
                Some(21.6),
                None,
                ["First header", "Second header", "Third header"]
                    .into_iter()
                    .take(line_count)
                    .map(|label| {
                        (
                            Alignment::Center,
                            true,
                            vec![hf_run(label, Some(family), size)],
                        )
                    })
                    .collect(),
            );
            let Page::Sheet(sheet) = &mut page else {
                unreachable!()
            };
            sheet.header = sheet.footer.take();
            let source = generate_typst(&make_doc(vec![page])).unwrap().source;
            let baselines: Vec<f64> = ["First", "Second", "Third"]
                .into_iter()
                .take(line_count)
                .map(|label| compiled_baseline_of(&source, label))
                .collect();
            for pair in baselines.windows(2) {
                assert!(
                    (pair[1] - pair[0] - advance).abs() < 0.01,
                    "{line_count}-line {family} {size}pt header: native advance {advance}pt, got {}",
                    pair[1] - pair[0],
                );
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn multiline_sheet_header_sections_keep_independent_first_lines() {
    for unequal_counts in [false, true] {
        let specifications = if unequal_counts {
            [
                (Alignment::Left, "Left", 11.0, 1),
                (Alignment::Center, "Center", 11.0, 3),
                (Alignment::Right, "Right", 11.0, 2),
            ]
        } else {
            [
                (Alignment::Left, "Left", 8.0, 2),
                (Alignment::Center, "Center", 24.0, 2),
                (Alignment::Right, "Right", 11.0, 2),
            ]
        };
        let mut page = sheet_page_with_footer_sections(
            PageSize {
                width: 612.0,
                height: 792.0,
            },
            72.0,
            Some(21.6),
            None,
            specifications
                .into_iter()
                .flat_map(|(alignment, label, size, count)| {
                    (0..count).map(move |index| {
                        (
                            alignment,
                            true,
                            vec![hf_run(
                                &format!("{label} line {index}"),
                                Some("Arial"),
                                size,
                            )],
                        )
                    })
                })
                .collect(),
        );
        let Page::Sheet(sheet) = &mut page else {
            unreachable!()
        };
        sheet.header = sheet.footer.take();
        let source = generate_typst(&make_doc(vec![page])).unwrap().source;
        let left = compiled_baseline_of(&source, "Left line 0");
        let center = compiled_baseline_of(&source, "Center line 0");
        let right = compiled_baseline_of(&source, "Right line 0");
        let (center_delta, right_delta) = if unequal_counts {
            (-1.0, -1.0)
        } else {
            (15.0, 3.0)
        };
        assert!(
            (center - left - center_delta).abs() < 0.01,
            "first center/left baselines {center}/{left}, expected delta {center_delta}"
        );
        assert!(
            (right - left - right_delta).abs() < 0.01,
            "first right/left baselines {right}/{left}, expected delta {right_delta}"
        );
    }
}

/// Native Excel footer probes match its wrapped-cell line advances (#1729).
/// Specs and measurements: tests/visual_audits/issue-1729/probes/.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn multiline_sheet_footer_uses_native_line_advance_and_keeps_last_seat() {
    for (family, size, advance) in [
        ("Arial", 8.0, 11.0),
        ("Arial", 11.0, 14.0),
        ("Arial", 12.0, 15.0),
        ("Arial", 14.0, 17.0),
        ("Arial", 24.0, 29.0),
        ("Aptos", 11.0, 14.0),
        ("Calibri", 11.0, 14.0),
        ("Malgun Gothic", 11.0, 17.0),
        ("Times New Roman", 11.0, 14.0),
    ] {
        let page = |labels: &[&str]| {
            sheet_page_with_footer_sections(
                PageSize {
                    width: 612.0,
                    height: 792.0,
                },
                72.0,
                Some(21.6),
                None,
                labels
                    .iter()
                    .map(|label| {
                        (
                            Alignment::Left,
                            true,
                            vec![hf_run(label, Some(family), size)],
                        )
                    })
                    .collect(),
            )
        };
        let single = generate_typst(&make_doc(vec![page(&["Last line"])]))
            .unwrap()
            .source;
        let multiple = generate_typst(&make_doc(vec![page(&["First line", "Last line"])]))
            .unwrap()
            .source;
        let first = compiled_baseline_of(&multiple, "First");
        let last = compiled_baseline_of(&multiple, "Last");
        assert!(
            (last - first - advance).abs() < 0.01,
            "{family} {size}pt: native advance {advance}, got {}",
            last - first
        );
        assert!(
            (last - compiled_baseline_of(&single, "Last")).abs() < 0.01,
            "{family} {size}pt: adding a line changed the last baseline"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn multiline_sheet_header_preserves_mixed_sizes_sections_and_fit_scale() {
    for (first_size, last_size, native_advance) in [
        (8.0, 11.0, 14.0),
        (24.0, 11.0, 17.0),
        (11.0, 8.0, 11.0),
        (11.0, 24.0, 26.0),
    ] {
        for scale in [1.0, 0.6, 0.8, 1.2] {
            let sections = [
                (Alignment::Left, "Left first", "Left last"),
                (Alignment::Center, "Center first", "Center last"),
                (Alignment::Right, "Right first", "Right last"),
            ];
            let paragraphs = sections
                .iter()
                .flat_map(|&(alignment, first, last)| {
                    [(first, first_size), (last, last_size)].map(|(label, size)| {
                        (
                            alignment,
                            true,
                            vec![hf_run(label, Some("Arial"), size * scale)],
                        )
                    })
                })
                .collect();
            let mut page = sheet_page_with_footer_sections(
                PageSize {
                    width: 612.0,
                    height: 792.0,
                },
                72.0,
                Some(21.6),
                Some(scale),
                paragraphs,
            );
            let Page::Sheet(sheet) = &mut page else {
                unreachable!()
            };
            sheet.header = sheet.footer.take();
            let source = generate_typst(&make_doc(vec![page])).unwrap().source;
            let runs = crate::render::pdf::compiled_text_runs(&source, 0).unwrap();
            for (_, first, last) in sections {
                let first_run = runs.iter().find(|run| run.text.contains(first)).unwrap();
                let last_run = runs.iter().find(|run| run.text.contains(last)).unwrap();
                let actual = last_run.baseline_pt - first_run.baseline_pt;
                assert!(
                    (actual - native_advance * scale).abs() < 0.02,
                    "{first_size}/{last_size} at {scale}: expected {}, got {actual}",
                    native_advance * scale
                );
            }
            let left = runs
                .iter()
                .find(|run| run.text.contains("Left last"))
                .unwrap();
            let center = runs
                .iter()
                .find(|run| run.text.contains("Center last"))
                .unwrap();
            let right = runs
                .iter()
                .find(|run| run.text.contains("Right last"))
                .unwrap();
            assert!(left.left_pt < center.left_pt && center.left_pt < right.left_pt);
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn multiline_sheet_footer_preserves_mixed_sizes_sections_and_fit_scale() {
    for (first_size, last_size, native_advance) in [
        (8.0, 11.0, 14.0),
        (24.0, 11.0, 17.0),
        (11.0, 8.0, 11.0),
        (11.0, 24.0, 26.0),
    ] {
        for scale in [1.0, 0.78] {
            let sections = [
                (Alignment::Left, "Left first", "Left last"),
                (Alignment::Center, "Center first", "Center last"),
                (Alignment::Right, "Right first", "Right last"),
            ];
            let paragraphs = sections
                .iter()
                .flat_map(|&(alignment, first, last)| {
                    [(first, first_size), (last, last_size)].map(|(label, size)| {
                        (
                            alignment,
                            true,
                            vec![hf_run(label, Some("Arial"), size * scale)],
                        )
                    })
                })
                .collect();
            let page = sheet_page_with_footer_sections(
                PageSize {
                    width: 612.0,
                    height: 792.0,
                },
                72.0,
                Some(21.6),
                Some(scale),
                paragraphs,
            );
            let source = generate_typst(&make_doc(vec![page])).unwrap().source;
            let runs = crate::render::pdf::compiled_text_runs(&source, 0).unwrap();
            for (_, first, last) in sections {
                let first_run = runs.iter().find(|run| run.text.contains(first)).unwrap();
                let last_run = runs.iter().find(|run| run.text.contains(last)).unwrap();
                let actual = last_run.baseline_pt - first_run.baseline_pt;
                assert!(
                    (actual - native_advance * scale).abs() < 0.02,
                    "{first_size}/{last_size} at {scale}: expected {}, got {actual}",
                    native_advance * scale
                );
            }
            let left = runs
                .iter()
                .find(|run| run.text.contains("Left last"))
                .unwrap();
            let center = runs
                .iter()
                .find(|run| run.text.contains("Center last"))
                .unwrap();
            let right = runs
                .iter()
                .find(|run| run.text.contains("Right last"))
                .unwrap();
            assert!(left.left_pt < center.left_pt && center.left_pt < right.left_pt);
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn multiline_sheet_footer_keeps_text_clear_of_inline_picture() {
    use crate::ir::{ImageData, ImageFormat};
    use crate::render::pdf::{PaintedKind, compiled_paint_sequence};
    let mut bitmap = std::io::Cursor::new(Vec::new());
    image::DynamicImage::new_rgb8(2, 2)
        .write_to(&mut bitmap, image::ImageFormat::Png)
        .unwrap();
    let mut page = sheet_page_with_footer_sections(
        PageSize {
            width: 612.0,
            height: 792.0,
        },
        100.0,
        Some(21.6),
        None,
        vec![
            (
                Alignment::Left,
                true,
                vec![hf_run("Prepared by Finance", Some("Arial"), 11.0)],
            ),
            (
                Alignment::Left,
                true,
                vec![hf_run("Approved", Some("Arial"), 11.0)],
            ),
        ],
    );
    let Page::Sheet(sheet) = &mut page else {
        unreachable!()
    };
    sheet.footer.as_mut().unwrap().paragraphs[1]
        .elements
        .push(HFInline::Image(ImageData {
            data: bitmap.into_inner(),
            format: ImageFormat::Png,
            width: Some(90.0),
            height: Some(40.0),
            rotation_deg: None,
            flip_h: false,
            flip_v: false,
            crop: None,
            stroke: None,
            alignment: None,
            clip_shape: None,
            shadow: None,
            paragraph_spacing: None,
        }));
    let output = generate_typst(&make_doc(vec![page])).unwrap();
    let painted = compiled_paint_sequence(&output.source, &output.images, 0).unwrap();
    let picture = painted
        .iter()
        .find(|item| item.kind == PaintedKind::Image)
        .unwrap();
    for text in painted.iter().filter(|item| item.kind == PaintedKind::Text) {
        let overlaps = text.bounds.0 < picture.bounds.2
            && text.bounds.2 > picture.bounds.0
            && text.bounds.1 < picture.bounds.3
            && text.bounds.3 > picture.bounds.1;
        assert!(
            !overlaps,
            "footer text overlaps its picture: {text:?} / {picture:?}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn multiline_sheet_header_counts_wrapped_lines_without_moving_first_line() {
    let long_line =
        "Prepared for the quarterly finance review and approved for internal distribution. "
            .repeat(2);
    for wraps_first in [true, false] {
        let labels = if wraps_first {
            [long_line.as_str(), "Footer second"]
        } else {
            ["Footer first", long_line.as_str()]
        };
        let mut page = sheet_page_with_footer_sections(
            PageSize {
                width: 612.0,
                height: 792.0,
            },
            72.0,
            Some(21.6),
            None,
            labels
                .iter()
                .map(|label| {
                    (
                        Alignment::Left,
                        true,
                        vec![hf_run(label, Some("Arial"), 11.0)],
                    )
                })
                .collect(),
        );
        let Page::Sheet(sheet) = &mut page else {
            unreachable!()
        };
        sheet.header = sheet.footer.take();
        let mut unwrapped = page.clone();
        let Page::Sheet(sheet) = &mut unwrapped else {
            unreachable!()
        };
        for paragraph in &mut sheet.header.as_mut().unwrap().paragraphs {
            paragraph.elements = vec![HFInline::Run(hf_run("Short header", Some("Arial"), 11.0))];
        }
        let unwrapped_source = generate_typst(&make_doc(vec![unwrapped])).unwrap().source;
        let expected_first = compiled_baseline_of(&unwrapped_source, "Short");
        let source = generate_typst(&make_doc(vec![page])).unwrap().source;
        let mut baselines: Vec<f64> = crate::render::pdf::compiled_text_runs(&source, 0)
            .unwrap()
            .iter()
            .filter(|run| run.text != "A")
            .map(|run| run.baseline_pt)
            .collect();
        baselines.sort_by(f64::total_cmp);
        baselines.dedup_by(|a, b| (*a - *b).abs() < 0.01);
        assert_eq!(baselines.len(), 3, "wrapped header lines: {baselines:?}");
        for (actual, expected) in
            baselines
                .iter()
                .zip([expected_first, expected_first + 14.0, expected_first + 28.0])
        {
            assert!(
                (actual - expected).abs() < 0.01,
                "native wrapped header baseline {expected}, got {baselines:?}; wraps_first={wraps_first}"
            );
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn multiline_sheet_footer_counts_wrapped_lines_in_native_advance() {
    let long_line =
        "Prepared for the quarterly finance review and approved for internal distribution. "
            .repeat(2);
    for wraps_first in [true, false] {
        let labels = if wraps_first {
            [long_line.as_str(), "Footer second"]
        } else {
            ["Footer first", long_line.as_str()]
        };
        let page = sheet_page_with_footer_sections(
            PageSize {
                width: 612.0,
                height: 792.0,
            },
            72.0,
            Some(21.6),
            None,
            labels
                .iter()
                .map(|label| {
                    (
                        Alignment::Left,
                        true,
                        vec![hf_run(label, Some("Arial"), 11.0)],
                    )
                })
                .collect(),
        );
        let source = generate_typst(&make_doc(vec![page])).unwrap().source;
        let mut baselines: Vec<f64> = crate::render::pdf::compiled_text_runs(&source, 0)
            .unwrap()
            .iter()
            .filter(|run| run.baseline_pt > 700.0)
            .map(|run| run.baseline_pt)
            .collect();
        baselines.sort_by(f64::total_cmp);
        baselines.dedup_by(|a, b| (*a - *b).abs() < 0.01);
        assert_eq!(baselines.len(), 3, "wrapped footer lines: {baselines:?}");
        for (actual, expected) in baselines.iter().zip([740.0, 754.0, 768.0]) {
            assert!(
                (actual - expected).abs() < 0.01,
                "native wrapped footer baseline {expected}, got {baselines:?}; wraps_first={wraps_first}"
            );
        }
    }
}

/// The bare `hhea` descent of `family` at `size_pt`, in points.
#[cfg(not(target_arch = "wasm32"))]
fn hhea_descent_pt(family: &str, size_pt: f64) -> f64 {
    let (_, descent_em, _) = crate::render::pdf::font_line_metrics_em(family)
        .unwrap_or_else(|| panic!("{family} metrics should resolve on every runner"));
    descent_em * size_pt
}

/// Baseline of the compiled page's run whose text contains `needle`, in
/// points down from the page top.
#[cfg(not(target_arch = "wasm32"))]
fn compiled_baseline_of(source: &str, needle: &str) -> f64 {
    let runs = crate::render::pdf::compiled_text_runs(source, 0)
        .unwrap_or_else(|error| panic!("compile failed: {error}\n{source}"));
    runs.iter()
        .find(|run| run.text.contains(needle))
        .unwrap_or_else(|| panic!("no run containing {needle:?}: {runs:?}\n{source}"))
        .baseline_pt
}

/// Excel lays a fitted sheet's header/footer out in sheet coordinates and
/// scales that box onto the paper. At 0.82, the A3 probe's 50pt page margins
/// therefore become the outward-rounded box 49.2..1141.44pt (#1510).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_fitted_sheet_footer_uses_its_scaled_horizontal_coordinate_box() {
    let mut page = sheet_page_with_seated_footer(Some(21.6), Some("Arial"));
    let Page::Sheet(sheet) = &mut page else {
        panic!("the fixture is a sheet page");
    };
    sheet.size = PageSize {
        width: 1191.0,
        height: 842.0,
    };
    sheet
        .footer
        .as_mut()
        .expect("the fixture has a footer")
        .sheet_print_scale = Some(0.82);

    let source = generate_typst(&make_doc(vec![page]))
        .expect("document should generate")
        .source;

    assert!(
        source.contains("#move(dx: -0.8pt)[#block(width: 1092.24pt)["),
        "the footer must use the scaled sheet-coordinate box: {source}"
    );
}

/// A seated sheet footer grows up from the page's bottom edge, not down from
/// the bottom margin (issue #1142).
///
/// Typst's default `footer-descent` is 30% of the bottom margin, so the band
/// moved with the body's own geometry: two pages of one workbook sharing
/// `<pageMargins bottom="0.75" footer="0.3"/>` put the same footer 5.21pt and
/// 7.13pt off the native export, and the gap differed between them. Pinning
/// the origin on the bottom margin line and spanning the remainder with a band
/// makes the seat depend on `@footer` alone.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_seated_sheet_footer_spans_the_gap_to_its_own_margin() {
    let source = generate_typst(&make_doc(vec![sheet_page_with_seated_footer(
        Some(21.6),
        Some("Arial"),
    )]))
    .expect("document should generate")
    .source;

    assert!(
        source.contains("footer-descent: 0pt"),
        "the footer origin must sit on the bottom margin line: {source}"
    );
    assert!(
        source.contains("block(width: 100%, height: 31pt)"),
        "the band must span the 54pt bottom margin down to the 23pt band Excel \
         leaves above a 0.3in (21.6pt, floored to 21pt) footer margin: {source}"
    );
    assert!(
        source.contains("#place(bottom"),
        "the content must rest on the band's bottom: {source}"
    );
}

/// Triangulation: the band is measured, not a constant (issue #1142).
///
/// The 2pt is measured: on Excel-for-Mac exports of one-factor variants of
/// `tests/fixtures/xlsx/headerFooterTest.xlsx`, a 12pt Calibri footer over a
/// 0.5in footer margin puts its baseline 41pt above the page's bottom edge —
/// 36pt of margin, Calibri's 3.22pt `hhea` descent, and 2pt between them. The
/// same series holds at 6, 8, 14, 20, 40 and 80pt, and across Arial, Verdana,
/// Times New Roman and Aptos.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_seated_sheet_footer_band_tracks_the_stated_seat() {
    let source = generate_typst(&make_doc(vec![sheet_page_with_seated_footer(
        Some(36.0),
        Some("Arial"),
    )]))
    .expect("document should generate")
    .source;

    assert!(
        source.contains("block(width: 100%, height: 16pt)"),
        "a 0.5in footer margin (36pt) plus the 2pt inset under a 54pt bottom \
         margin leaves a 16pt band: {source}"
    );
}

/// The band states where the footer's baseline lands, in points above the
/// band's bottom: the face's own `hhea` descent, rounded with the band to the
/// whole point Excel prints (issues #1142, #1552).
///
/// Arial, because every runner resolves it — through Liberation Sans where the
/// face itself is absent.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_seated_sheet_footer_states_its_own_bottom_edge() {
    let descent_pt: f64 = hhea_descent_pt("Arial", 8.0);
    let expected_lift_pt: f64 = (23.0 + descent_pt).round() - 23.0;
    assert!(
        (1.0..2.5).contains(&descent_pt),
        "the test needs a descent that rounds visibly: Arial 8pt gives {descent_pt}pt"
    );

    let source = generate_typst(&make_doc(vec![sheet_page_with_seated_footer(
        Some(21.6),
        Some("Arial"),
    )]))
    .expect("document should generate")
    .source;

    assert!(
        !source.contains("bottom-edge: \"descender\"") && !source.contains("em)"),
        "the band must state the seat in points, not a normalised or em descent: {source}"
    );
    assert!(
        source.contains(&format!("bottom-edge: -{}pt", format_f64(expected_lift_pt))),
        "the baseline must sit round(23 + {descent_pt}) - 23 = {expected_lift_pt}pt above \
         the band: {source}"
    );
}

/// A footer with no seat keeps the old placement, so a `Page::Sheet` built
/// without one is untouched.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn an_unseated_sheet_footer_keeps_the_default_descent() {
    let source = generate_typst(&make_doc(vec![sheet_page_with_seated_footer(None, None)]))
        .expect("document should generate")
        .source;

    assert!(
        !source.contains("footer-descent"),
        "an unseated footer must not pin the origin: {source}"
    );
}

/// Excel seats a printed footer line on a whole point: the band, plus the
/// deepest run's `hhea` descent, rounded (issue #1552).
///
/// Native Excel-for-Mac exports of `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`
/// with its 8pt Aptos label reset to 10, 16, 20, 24 and 40pt seat the baseline
/// 25, 27, 28, 29 and 33pt above the paper, each `round(23 + descent) - 1`;
/// the one-point drop is the rich-text path below. The reported label — a
/// 1pt `#` in the Normal font ahead of the 8pt run — had taken the 1pt run's
/// face for the 8pt run's descent and landed at 24.78pt against Excel's 24.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_mixed_size_sheet_footer_seats_its_deepest_run_on_a_whole_point() {
    const PAGE_HEIGHT_PT: f64 = 792.0;
    let page = sheet_page_with_footer_sections(
        PageSize {
            width: 612.0,
            height: PAGE_HEIGHT_PT,
        },
        54.0,
        Some(21.6),
        None,
        vec![(
            Alignment::Left,
            true,
            vec![
                hf_run("#", Some("Libertinus Serif"), 1.0),
                hf_run(" Sensitivity: Internal", Some("Arial"), 8.0),
            ],
        )],
    );
    let deepest_pt: f64 =
        hhea_descent_pt("Libertinus Serif", 1.0).max(hhea_descent_pt("Arial", 8.0));
    let expected_seat_pt: f64 = (23.0 + deepest_pt).round() - 1.0;

    let source = generate_typst(&make_doc(vec![page]))
        .expect("document should generate")
        .source;
    let baseline_pt: f64 = compiled_baseline_of(&source, "Sensitivity");

    assert!(
        (PAGE_HEIGHT_PT - baseline_pt - expected_seat_pt).abs() < 0.05,
        "the 8pt run must sit {expected_seat_pt}pt above the paper, got {:.3}pt\n{source}",
        PAGE_HEIGHT_PT - baseline_pt
    );
    assert!(
        (compiled_baseline_of(&source, "#") - baseline_pt).abs() < 0.01,
        "both runs share the line's baseline\n{source}"
    );
}

/// The deepest run decides the line, not the first one and not the largest
/// one: in native exports a 20pt `#` ahead of 8pt text seats on the 20pt
/// run's descent, and 13pt Aptos beside 15pt Arial on the 13pt Aptos's 3.66pt
/// rather than the larger run's 3.18pt, in either order (issue #1552).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn the_deepest_sheet_footer_run_wins_in_either_order() {
    const PAGE_HEIGHT_PT: f64 = 792.0;
    let big = || hf_run("#", Some("Arial"), 20.0);
    let small = || hf_run(" Sensitivity: Internal", Some("Arial"), 8.0);
    let expected_seat_pt: f64 = (23.0 + hhea_descent_pt("Arial", 20.0)).round() - 1.0;

    for runs in [vec![big(), small()], vec![small(), big()]] {
        let page = sheet_page_with_footer_sections(
            PageSize {
                width: 612.0,
                height: PAGE_HEIGHT_PT,
            },
            54.0,
            Some(21.6),
            None,
            vec![(Alignment::Left, true, runs)],
        );
        let source = generate_typst(&make_doc(vec![page]))
            .expect("document should generate")
            .source;
        let seat_pt: f64 = PAGE_HEIGHT_PT - compiled_baseline_of(&source, "Sensitivity");
        assert!(
            (seat_pt - expected_seat_pt).abs() < 0.05,
            "the 20pt run's descent seats the line at {expected_seat_pt}pt whichever \
             side it is on, got {seat_pt:.3}pt\n{source}"
        );
    }
}

/// A section drawn as one uniform run sits one point higher than the same
/// text on the rich-text path, and the two paths are decided per section:
/// native Excel seats `&L_x000D_&1#&"Aptos"&8&K000000 Sensitivity: Internal&R&"Aptos"&8Page`
/// with the left label 24pt and the right one 25pt above the paper
/// (issue #1552).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_plain_sheet_footer_section_sits_one_point_above_a_rich_one() {
    const PAGE_HEIGHT_PT: f64 = 792.0;
    let page = sheet_page_with_footer_sections(
        PageSize {
            width: 612.0,
            height: PAGE_HEIGHT_PT,
        },
        54.0,
        Some(21.6),
        None,
        vec![
            (
                Alignment::Left,
                true,
                vec![
                    hf_run("#", Some("Arial"), 1.0),
                    hf_run(" Sensitivity: Internal", Some("Arial"), 8.0),
                ],
            ),
            (
                Alignment::Right,
                false,
                vec![hf_run("Page", Some("Arial"), 8.0)],
            ),
        ],
    );
    let plain_seat_pt: f64 = (23.0 + hhea_descent_pt("Arial", 8.0)).round();

    let source = generate_typst(&make_doc(vec![page]))
        .expect("document should generate")
        .source;
    let rich_baseline_pt: f64 = compiled_baseline_of(&source, "Sensitivity");
    let plain_baseline_pt: f64 = compiled_baseline_of(&source, "Page");

    assert!(
        (PAGE_HEIGHT_PT - plain_baseline_pt - plain_seat_pt).abs() < 0.05,
        "the plain section sits on round(23 + descent) = {plain_seat_pt}pt, got {:.3}pt\n{source}",
        PAGE_HEIGHT_PT - plain_baseline_pt
    );
    assert!(
        (rich_baseline_pt - plain_baseline_pt - 1.0).abs() < 0.05,
        "the rich section sits exactly one point lower than the plain one, got {:.3}pt\n{source}",
        rich_baseline_pt - plain_baseline_pt
    );
}

/// A fitted sheet seats its footer in whole *sheet* points and scales the
/// result onto the paper: at 0.78 the A3 budget sheet's 0.3in margin floors
/// to 27 sheet points, and native exports of its 8, 16, 24 and 40pt labels
/// land at (30, 33, 35, 39) x 0.78pt above the paper — `round(29 + descent) - 1`
/// each, with a plain 8pt run at 31 x 0.78 (issue #1552).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_fitted_sheet_footer_seat_is_a_whole_sheet_point_scaled_to_paper() {
    const PAGE_HEIGHT_PT: f64 = 1191.0;
    const SCALE: f64 = 0.78;
    for (is_rich, drop_pt) in [(true, 1.0), (false, 0.0)] {
        let page = sheet_page_with_footer_sections(
            PageSize {
                width: 842.0,
                height: PAGE_HEIGHT_PT,
            },
            54.0,
            Some(21.6),
            Some(SCALE),
            vec![(
                Alignment::Left,
                is_rich,
                // The parser has already multiplied the run's size by the scale.
                vec![hf_run(" Sensitivity: Internal", Some("Arial"), 8.0 * SCALE)],
            )],
        );
        let band_sheet_pt: f64 = (21.6 / SCALE).floor() + 2.0;
        let seat_sheet_pt: f64 = (band_sheet_pt + hhea_descent_pt("Arial", 8.0)).round() - drop_pt;
        let expected_seat_pt: f64 = seat_sheet_pt * SCALE;

        let source = generate_typst(&make_doc(vec![page]))
            .expect("document should generate")
            .source;
        let seat_pt: f64 = PAGE_HEIGHT_PT - compiled_baseline_of(&source, "Sensitivity");
        assert!(
            (seat_pt - expected_seat_pt).abs() < 0.05,
            "rich={is_rich}: the seat is {seat_sheet_pt} sheet points x {SCALE} = \
             {expected_seat_pt:.2}pt, got {seat_pt:.3}pt\n{source}"
        );
    }
}

#[test]
fn test_table_page_no_header_footer() {
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: make_simple_table(vec![vec!["A"]]),
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.contains("header:"));
    assert!(!output.source.contains("footer:"));
}

/// The single-series `Sales` bar chart the anchored-chart render tests float
/// over a small grid.
fn sales_bar_chart() -> crate::ir::Chart {
    use crate::ir::{Chart, ChartGrouping, ChartSeries, ChartType, DataLabels, LegendPosition};

    Chart {
        chart_type: ChartType::Bar,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Sales".to_string()),
        categories: vec!["Q1".to_string(), "Q2".to_string()],
        series: vec![ChartSeries {
            name: Some("Revenue".to_string()),
            values: vec![100.0, 200.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

#[test]
fn test_table_page_with_anchored_chart_overlays_the_grid() {
    let chart = sales_bar_chart();

    let page = Page::Sheet(SheetPage {
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
        charts: vec![crate::ir::SheetChart {
            anchor_row: 3,
            placement: Some(crate::ir::SheetChartPlacement {
                x_offset_pt: 40.0,
                y_offset_pt: 60.0,
                width: 200.0,
                height: 120.0,
                print_scale: 1.0,
                clip_window: None,
            }),
            chart,
        }],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });

    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    let src = &output.source;

    // Excel floats an anchored chart over the cells, so the grid keeps every
    // row in one table instead of being cut into segments around it (#982).
    assert_eq!(src.matches("#table(").count(), 1);
    // The anchor's offsets are the sheet content origin's, and the drawing
    // layer measures from the page corner, so each carries its margin.
    let margin: f64 = crate::defaults::DEFAULT_MARGIN_PT;
    let placement: String = format!(
        "#place(top + left, dy: {}pt)[#place(top + left, dx: {}pt)[",
        margin + 60.0,
        margin + 40.0
    );
    let overlay: usize = src
        .find(&placement)
        .unwrap_or_else(|| panic!("the chart is placed at its anchor's offsets: {src}"));
    // The drawings float in the page foreground, which paints above the whole
    // body however early in the source it is declared (issue #1168).
    let foreground: usize = src
        .find(", foreground: ")
        .expect("the drawing layer is the sheet's page foreground");
    let chart_pos: usize = src.find("Sales").expect("the chart's title");
    let table_pos: usize = src.find("#table(").expect("the sheet grid");
    assert!(
        foreground < overlay && overlay < chart_pos && chart_pos < table_pos,
        "the chart is drawn in the page foreground the grid follows"
    );
    // The anchor sizes the chart, the way a slide's graphicFrame extent does:
    // its title band and plot box together fill the anchored 200x120pt.
    assert!(
        src.contains("#block(width: 200pt, height: 19pt")
            && src.contains("#box(width: 200pt, height: 101pt"),
        "the chart is laid out at the anchor's size"
    );
}

#[test]
fn test_a_chart_continued_onto_a_page_column_is_clipped_to_its_window() {
    // Width pagination handed this page the second tile of a chart that
    // crosses the boundary: shifted 100pt left of the tile's content edge
    // and clipped to the tile's 300pt window (issue #1598).
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize {
            width: 500.0,
            height: 800.0,
        },
        margins: Margins::default(),
        table: make_simple_table(vec![vec!["Row 1"], vec!["Row 2"]]),
        header: None,
        footer: None,
        charts: vec![crate::ir::SheetChart {
            anchor_row: 1,
            placement: Some(crate::ir::SheetChartPlacement {
                x_offset_pt: -100.0,
                y_offset_pt: 60.0,
                width: 200.0,
                height: 120.0,
                print_scale: 1.0,
                clip_window: Some(crate::ir::SheetClipWindow {
                    left_pt: 0.0,
                    width_pt: 300.0,
                }),
            }),
            chart: sales_bar_chart(),
        }],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });

    let output = generate_typst(&make_doc(vec![page])).unwrap();
    let src = &output.source;
    let margin: f64 = crate::defaults::DEFAULT_MARGIN_PT;
    // The clip box stands at the window's page position and reaches the page
    // bottom, so only the horizontal edges bound the chart; the chart itself
    // is placed inside it at its negative offset.
    let clipped: String = format!(
        "#place(top + left, dy: {}pt)[#place(top + left, dx: {}pt)[#box(width: 300pt, height: 800pt, clip: true)[#place(top + left, dx: -100pt)[",
        margin + 60.0,
        margin
    );
    assert!(
        src.contains(&clipped),
        "the continued chart is clipped to its page-column window: {src}"
    );
    assert!(
        src.contains("#block(width: 200pt, height: 19pt")
            && src.contains("#box(width: 200pt, height: 101pt"),
        "the clipped chart is still laid out at its anchor's full size"
    );
}

#[test]
fn test_a_chart_continued_after_repeated_title_columns_starts_its_window_there() {
    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize {
            width: 500.0,
            height: 800.0,
        },
        margins: Margins::default(),
        table: make_simple_table(vec![vec!["Row 1"], vec!["Row 2"]]),
        header: None,
        footer: None,
        charts: vec![crate::ir::SheetChart {
            anchor_row: 1,
            placement: Some(crate::ir::SheetChartPlacement {
                x_offset_pt: 40.0,
                y_offset_pt: 60.0,
                width: 200.0,
                height: 120.0,
                print_scale: 1.0,
                clip_window: Some(crate::ir::SheetClipWindow {
                    left_pt: 80.0,
                    width_pt: 300.0,
                }),
            }),
            chart: sales_bar_chart(),
        }],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });

    let output = generate_typst(&make_doc(vec![page])).unwrap();
    let src = &output.source;
    let margin: f64 = crate::defaults::DEFAULT_MARGIN_PT;
    // The window begins after the 80pt of repeated titles, and the chart's
    // offset is measured from that edge, so the inner place is 40pt in.
    let clipped: String = format!(
        "#place(top + left, dx: {}pt)[#box(width: 300pt, height: 800pt, clip: true)[#place(top + left, dx: -40pt)[",
        margin + 80.0
    );
    assert!(
        src.contains(&clipped),
        "the window starts after the repeated title columns: {src}"
    );
}

/// Excel scales a printed sheet whole, drawings included, so a fit-to-page
/// scale shrinks a chart's text, tick marks and legend along with its frame.
/// Shrinking the frame alone left the reported workbook's tick labels and
/// legend entries at the size the chart XML declares, about 22% larger than
/// Excel prints them (issue #1069).
#[test]
fn test_fitted_sheet_draws_the_whole_chart_shrunk() {
    let unscaled: String = sheet_source_with_chart_print_scale(1.0);
    let fitted: String = sheet_source_with_chart_print_scale(0.82);

    assert!(
        !unscaled.contains("#scale("),
        "a sheet printed at full size wraps the chart in no transform"
    );
    let wrapper: &str = "#scale(x: 82%, y: 82%, origin: top + left)[";
    let dx_pt: f64 = crate::defaults::DEFAULT_MARGIN_PT + 40.0;
    assert!(
        fitted.contains(&format!("#place(top + left, dx: {dx_pt}pt)[{wrapper}")),
        "the fitted sheet shrinks the drawing from the anchor's top-left corner"
    );
    // The chart still lays itself out at the anchor's full frame and its own
    // type sizes; the transform is what shrinks it, so the text, tick marks,
    // legend and plot all come down by the same factor.
    assert!(
        fitted.contains("#block(width: 200pt, height: 19pt")
            && fitted.contains("#box(width: 200pt, height: 101pt"),
        "the fitted chart is laid out at the anchor's full size"
    );
    assert_eq!(
        fitted.len(),
        unscaled.len() + wrapper.len() + "]".len(),
        "the transform and its closing bracket are the only difference"
    );
    // The transform is markup the layout engine has to accept: a source-only
    // assertion would pass on a `#scale` argument Typst rejects, and every
    // fitted sheet carrying a chart would fail to convert at all.
    crate::render::pdf::compile_to_pdf(&fitted, &[], None, &[], false, false)
        .expect("the fitted sheet compiles");
}

/// The sheet of [`test_table_page_with_anchored_chart_overlays_the_grid`],
/// printed at `print_scale`.
fn sheet_source_with_chart_print_scale(print_scale: f64) -> String {
    generate_typst(&make_doc(vec![Page::Sheet(
        sheet_page_with_chart_print_scale(print_scale),
    )]))
    .unwrap()
    .source
}

fn sheet_page_with_chart_print_scale(print_scale: f64) -> SheetPage {
    use crate::ir::{Chart, ChartGrouping, ChartSeries, ChartType, DataLabels, LegendPosition};

    let chart = Chart {
        chart_type: ChartType::Bar,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Sales".to_string()),
        categories: vec!["Q1".to_string(), "Q2".to_string()],
        series: vec![ChartSeries {
            name: Some("Revenue".to_string()),
            values: vec![100.0, 200.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    };

    SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: make_simple_table(vec![vec!["Row 1"], vec!["Row 2"]]),
        header: None,
        footer: None,
        charts: vec![crate::ir::SheetChart {
            anchor_row: 3,
            placement: Some(crate::ir::SheetChartPlacement {
                x_offset_pt: 40.0,
                y_offset_pt: 60.0,
                width: 200.0,
                height: 120.0,
                print_scale,
                clip_window: None,
            }),
            chart,
        }],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    }
}

#[test]
fn test_table_page_with_chart_at_end() {
    use crate::ir::{Chart, ChartGrouping, ChartSeries, ChartType, DataLabels, LegendPosition};

    let chart = Chart {
        chart_type: ChartType::Pie,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Pie".to_string()),
        categories: vec!["A".to_string()],
        series: vec![ChartSeries {
            name: None,
            values: vec![100.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    };

    let page = Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: make_simple_table(vec![vec!["Data"]]),
        header: None,
        footer: None,
        charts: vec![crate::ir::SheetChart {
            anchor_row: u32::MAX,
            placement: None,
            chart,
        }],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    });

    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    let src = &output.source;

    let table_pos = src.find("#table(").unwrap();
    let chart_pos = src.find("Pie").unwrap();
    assert!(table_pos < chart_pos);
}

#[test]
fn test_paper_size_override_letter() {
    use crate::config::PaperSize;

    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
    let options = ConvertOptions {
        paper_size: Some(PaperSize::Letter),
        ..Default::default()
    };
    let output = generate_typst_with_options(&doc, &options).unwrap();
    assert!(output.source.contains("width: 612pt"));
    assert!(output.source.contains("height: 792pt"));
}

#[test]
fn test_landscape_override_swaps_dimensions() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
    let options = ConvertOptions {
        landscape: Some(true),
        ..Default::default()
    };
    let output = generate_typst_with_options(&doc, &options).unwrap();
    assert!(output.source.contains("width: 841.89pt"));
    assert!(output.source.contains("height: 595.28pt"));
}

#[test]
fn test_portrait_override_keeps_portrait() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
    let options = ConvertOptions {
        landscape: Some(false),
        ..Default::default()
    };
    let output = generate_typst_with_options(&doc, &options).unwrap();
    assert!(output.source.contains("width: 595.28pt"));
    assert!(output.source.contains("height: 841.89pt"));
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
    assert!(output.source.contains("width: 792pt"));
    assert!(output.source.contains("height: 612pt"));
}

#[test]
fn test_no_override_uses_original_size() {
    let doc = make_doc(vec![make_flow_page(vec![make_paragraph("Test")])]);
    let options = ConvertOptions::default();
    let output = generate_typst_with_options(&doc, &options).unwrap();
    assert!(output.source.contains("width: 595.28pt"));
}

/// Word letterhead headers commonly carry a `w:pBdr/w:bottom` rule under the
/// header text. The rule must render below the content, not be dropped.
#[test]
fn test_generate_header_with_bottom_border_draws_rule_below_text() {
    use crate::ir::{BorderSide, CellBorder, HFInline, HeaderFooter, HeaderFooterParagraph};

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Manual v0.6".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: Some(CellBorder {
                    top: None,
                    bottom: Some(BorderSide {
                        width: 0.5,
                        color: Color::new(0xCC, 0xCC, 0xCC),
                        style: BorderLineStyle::Solid,
                        join: LineJoin::Round,
                    }),
                    left: None,
                    right: None,
                }),
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert_eq!(
        output.source.matches("line(length: 100%").count(),
        1,
        "the bottom rule must be emitted"
    );
    assert!(
        output.source.contains("rgb(204, 204, 204)"),
        "the rule keeps the pBdr color"
    );
    let text_pos = output
        .source
        .find("Manual v0.6")
        .expect("header text present");
    let rule_pos = output
        .source
        .find("line(length: 100%")
        .expect("rule present");
    assert!(
        rule_pos > text_pos,
        "a bottom rule must be drawn after the header text"
    );
    // The rule overhangs the text column, so it must not inherit the header
    // paragraph's alignment — a right-aligned header would otherwise pin the
    // line's right edge to the column and throw the whole overhang left
    // (issue #840).
    assert!(
        output
            .source
            .contains("#align(left)[#move(dx: -1.44pt)[#line(length: 100% + 2.88pt"),
        "the rule states its own alignment: {}",
        output.source
    );
}

/// A right-aligned header still draws its rule symmetrically about the column.
///
/// Regression for #840: the rule is 2.88pt wider than the text column, so
/// inheriting `w:jc = right` put its right edge on the column edge and the
/// whole overhang on the left, which the `#move` then doubled.
#[test]
fn a_right_aligned_header_does_not_drag_its_rule_left() {
    use crate::ir::{BorderSide, CellBorder, HFInline, HeaderFooter, HeaderFooterParagraph};

    let right_aligned = ParagraphStyle {
        alignment: Some(crate::ir::Alignment::Right),
        ..ParagraphStyle::default()
    };
    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: right_aligned,
                elements: vec![HFInline::Run(Run {
                    text: "Minutes | Internal".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: Some(CellBorder {
                    top: None,
                    bottom: Some(BorderSide {
                        width: 0.5,
                        color: Color::new(0xCC, 0xCC, 0xCC),
                        style: BorderLineStyle::Solid,
                        join: LineJoin::Round,
                    }),
                    left: None,
                    right: None,
                }),
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let source = generate_typst(&doc).unwrap().source;
    assert!(
        source.contains("#align(left)[#move(dx: -1.44pt)[#line(length: 100% + 2.88pt"),
        "the rule must override the paragraph's right alignment: {source}"
    );
}

/// A paragraph carrying both a top and a bottom rule must draw both.
#[test]
fn test_generate_header_with_top_and_bottom_borders_draws_both_rules() {
    use crate::ir::{BorderSide, CellBorder, HFInline, HeaderFooter, HeaderFooterParagraph};

    let rule = |width: f64| {
        Some(BorderSide {
            width,
            color: Color::new(0x33, 0x66, 0x99),
            style: BorderLineStyle::Solid,
            join: LineJoin::Round,
        })
    };

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Framed".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: Some(CellBorder {
                    top: rule(1.0),
                    bottom: rule(1.0),
                    left: None,
                    right: None,
                }),
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert_eq!(
        output.source.matches("line(length: 100%").count(),
        2,
        "both rules must be emitted"
    );
}

/// Word measures `w:pgMar/@w:footer` from the bottom page edge to the bottom of
/// the footer, so the footer must be pinned to that line and grow upward — not
/// pushed further away from the edge.
#[test]
fn test_flow_page_footer_is_pinned_to_the_word_edge_distance() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins {
            top: 62.35,
            bottom: 62.35,
            left: 70.85,
            right: 70.85,
        },
        content: vec![make_paragraph("Body")],
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: Some(35.4),
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "- 1 -".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("footer-descent: 0pt"),
        "the footer origin must sit on the bottom margin line"
    );
    assert!(
        output
            .source
            .contains("block(width: 100%, height: 26.95pt)"),
        "the band must span bottom margin minus footer distance, got: {}",
        output.source
    );
    assert!(
        output.source.contains("bottom-edge: \"descender\""),
        "Word measures to the descender line"
    );
    assert!(
        output.source.contains("place(bottom"),
        "the footer grows upward from the pinned bottom"
    );
    assert!(
        !output.source.contains("move(dy: -26.95pt)"),
        "the footer must not be shifted away from the page edge"
    );
}

/// Without a declared footer distance the previous placement is kept, so
/// formats that do not carry the attribute are unaffected.
#[test]
fn test_flow_page_footer_without_edge_distance_keeps_default_placement() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Plain footer".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("footer: ["));
    assert!(!output.source.contains("footer-descent"));
}

/// A footer distance at or beyond the bottom margin leaves no band to draw, so
/// the default placement is used instead of a zero or negative height block.
#[test]
fn test_flow_page_footer_distance_beyond_margin_falls_back() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins {
            top: 72.0,
            bottom: 30.0,
            left: 72.0,
            right: 72.0,
        },
        content: vec![make_paragraph("Body")],
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: Some(48.0),
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Deep footer".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.contains("footer-descent"));
    assert!(output.source.contains("Deep footer"));
}

/// Word keeps `w:spacing w:before` on the very first body paragraph, while
/// Typst drops leading block spacing at a page boundary. The gap has to be
/// emitted as explicit vertical space so the first heading is not pulled up to
/// the top margin.
#[test]
fn test_first_document_paragraph_keeps_its_space_before() {
    let mut heading = make_paragraph("Research Report");
    if let Block::Paragraph(ref mut paragraph) = heading {
        paragraph.style.space_before = Some(14.0);
        paragraph.style.space_after = Some(7.0);
    }

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![heading, make_paragraph("Body")],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    let spacer = output
        .source
        .find("#v(14pt")
        .expect("explicit leading space emitted");
    let heading_pos = output
        .source
        .find("Research Report")
        .expect("heading present");
    assert!(
        spacer < heading_pos,
        "the space must precede the heading, got: {}",
        output.source
    );
    assert!(
        !output.source[spacer..heading_pos].contains("above: 14pt"),
        "the collapsed block spacing must not be emitted twice"
    );
}

/// Only the document's first paragraph gets the explicit gap. Word suppresses
/// space-before at the top of a page reached by a break, so later paragraphs
/// keep ordinary collapsing block spacing.
#[test]
fn test_later_paragraph_space_before_stays_block_spacing() {
    let mut second = make_paragraph("Second");
    if let Block::Paragraph(ref mut paragraph) = second {
        paragraph.style.space_before = Some(21.0);
    }

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("First"), second],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        !output.source.contains("#v(21pt"),
        "later paragraphs keep block spacing"
    );
    assert!(output.source.contains("above: 21pt"));
}

/// `w:pBdr` sides declare a `w:space` gap in points between the text and the
/// rule; a header rule must sit that far below its text.
#[test]
fn test_generate_header_border_uses_declared_pbdr_space() {
    use crate::ir::{
        BorderSide, CellBorder, HFInline, HeaderFooter, HeaderFooterParagraph, Insets,
    };

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Manual v0.6".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: Some(CellBorder {
                    top: None,
                    bottom: Some(BorderSide {
                        width: 0.5,
                        color: Color::new(0xCC, 0xCC, 0xCC),
                        style: BorderLineStyle::Solid,
                        join: LineJoin::Round,
                    }),
                    left: None,
                    right: None,
                }),
                border_space: Some(Insets {
                    top: 0.0,
                    right: 0.0,
                    bottom: 4.0,
                    left: 0.0,
                }),
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("block(height: 4pt)[]"),
        "the declared 4pt gap must separate text and rule, got: {}",
        output.source
    );
    assert!(
        output.source.contains("bottom-edge: \"descender\""),
        "Word measures the gap from the descender line"
    );
}

/// Without `w:space` the rule keeps the previous hairline clearance.
#[test]
fn test_generate_header_border_without_space_keeps_hairline_gap() {
    use crate::ir::{BorderSide, CellBorder, HFInline, HeaderFooter, HeaderFooterParagraph};

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Plain".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: Some(CellBorder {
                    top: None,
                    bottom: Some(BorderSide {
                        width: 0.5,
                        color: Color::black(),
                        style: BorderLineStyle::Solid,
                        join: LineJoin::Round,
                    }),
                    left: None,
                    right: None,
                }),
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("block(height: 0.5pt)[]"));
}

/// Word measures `w:pgMar/@w:header` from the top page edge to the top of the
/// header, which then grows downward. Typst anchors headers by their bottom, so
/// the band has to hold the content against the header top.
#[test]
fn test_flow_page_header_is_pinned_to_the_word_edge_distance() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins {
            top: 62.35,
            bottom: 62.35,
            left: 70.85,
            right: 70.85,
        },
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: Some(35.4),
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Manual v0.6".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("header-ascent: 0pt"),
        "the header origin must sit on the top margin line"
    );
    assert!(
        output
            .source
            .contains("block(width: 100%, height: 26.95pt)"),
        "the band must span top margin minus header distance, got: {}",
        output.source
    );
    assert!(
        output.source.contains("place(top"),
        "the header grows downward from the pinned top"
    );
}

/// Without a declared header distance the previous placement is kept.
#[test]
fn test_flow_page_header_without_edge_distance_keeps_default_placement() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Plain header".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(output.source.contains("header: ["));
    assert!(!output.source.contains("header-ascent"));
    assert!(
        !output.source.contains("place(top, dy:"),
        "an unpinned header has no origin to seat a baseline against: {}",
        output.source
    );
}

/// The faces that carry Arial's metrics: the corpus baselines below are Arial's,
/// and Liberation Sans and Arimo are metric-compatible clones of it. When the
/// machine has none of them the substitute chain lands somewhere else and the
/// absolute numbers no longer apply.
const ARIAL_METRIC_FACES: [&str; 3] = ["Arial", "Liberation Sans", "Arimo"];

/// Malgun Gothic has no metric-compatible substitute, so only the face itself
/// can be held to the Korean corpus baselines.
const MALGUN_METRIC_FACES: [&str; 1] = ["Malgun Gothic"];

/// Build a one-section document whose header holds the given paragraphs.
fn doc_with_header(
    header_distance_pt: Option<f64>,
    top_margin_pt: f64,
    paragraphs: Vec<crate::ir::HeaderFooterParagraph>,
) -> Document {
    use crate::ir::HeaderFooter;

    make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins {
            top: top_margin_pt,
            bottom: 62.35,
            left: 70.85,
            right: 70.85,
        },
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: header_distance_pt,
            sheet_print_scale: None,
            paragraphs,
        }),
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })])
}

/// A header paragraph holding one styled run.
fn header_text_paragraph(text: &str, style: TextStyle) -> crate::ir::HeaderFooterParagraph {
    use crate::ir::{HFInline, HeaderFooterParagraph};

    HeaderFooterParagraph {
        style: ParagraphStyle::default(),
        elements: vec![HFInline::Run(Run {
            text: text.to_string(),
            style,
            href: None,
            footnote: None,
        })],
        border: None,
        border_space: None,
        sheet_section_is_rich: false,
        frame: None,
    }
}

fn arial(size_pt: f64) -> TextStyle {
    TextStyle {
        font_family: Some("Arial".to_string()),
        font_size: Some(size_pt),
        ..TextStyle::default()
    }
}

/// Build a one-section document whose header holds a single run.
fn doc_with_header_run(
    header_distance_pt: Option<f64>,
    top_margin_pt: f64,
    text: &str,
    style: TextStyle,
) -> Document {
    doc_with_header(
        header_distance_pt,
        top_margin_pt,
        vec![header_text_paragraph(text, style)],
    )
}

#[cfg(not(target_arch = "wasm32"))]
/// Compile the document and report every text run the layout engine placed on
/// page 1, ordered down the page.
///
/// Placement assertions read these rather than the emitted source: `place`,
/// `measure` and `top-edge` are all resolved by the layout engine, and the
/// first attempt at this placement passed every source assertion while moving
/// each wrapped line of the paragraph it touched (issue #629).
fn placed_runs(doc: &Document) -> Vec<crate::render::pdf::PlacedTextRun> {
    let output = generate_typst(doc).expect("document should generate");
    let mut runs = crate::render::pdf::compiled_text_runs(&output.source, 0)
        .unwrap_or_else(|error| panic!("compile failed: {error}\n{}", output.source));
    runs.sort_by(|left, right| {
        left.baseline_pt
            .total_cmp(&right.baseline_pt)
            .then(left.left_pt.total_cmp(&right.left_pt))
    });
    runs
}

#[cfg(not(target_arch = "wasm32"))]
/// Whether the first run carrying `needle` was shaped by one of `families`.
///
/// The corpus baselines are properties of a specific face, and the font chain
/// silently substitutes: asserting them against whatever the machine happens to
/// have installed would test the substitute, not the placement.
fn shaped_by(doc: &Document, needle: &str, families: &[&str]) -> bool {
    placed_runs(doc)
        .iter()
        .find(|run| run.text.contains(needle))
        .is_some_and(|run| {
            families
                .iter()
                .any(|family| run.family.eq_ignore_ascii_case(family))
        })
}

#[cfg(not(target_arch = "wasm32"))]
/// The distinct baselines, top to bottom, of the runs whose text contains
/// `needle`. Wrapped lines of one paragraph each contribute one entry.
fn baselines_of(doc: &Document, needle: &str) -> Vec<f64> {
    let mut baselines: Vec<f64> = Vec::new();
    for run in placed_runs(doc) {
        if !run.text.contains(needle) {
            continue;
        }
        if baselines
            .last()
            .is_none_or(|last: &f64| (run.baseline_pt - last).abs() > 0.01)
        {
            baselines.push(run.baseline_pt);
        }
    }
    baselines
}

#[cfg(not(target_arch = "wasm32"))]
/// Word seats the header's first baseline one font ascent below
/// `w:pgMar/@w:header`, not at a proportion of the top margin.
///
/// `05_technical_manual_en` declares `w:top="1247" w:header="708"` — 62.35pt and
/// 35.40pt — over an 8pt Arial run, and its native export puts that baseline at
/// 42.72pt on the 0.24pt grid Word quantises to, against `35.40 + 0.9053 x 8 =
/// 42.64` predicted (issue #629).
#[test]
fn test_header_first_baseline_sits_one_font_ascent_below_the_header_distance() {
    let ascender_em: f64 =
        crate::render::pdf::font_hhea_ascender_em("Arial").expect("Arial metrics should resolve");
    let doc = doc_with_header_run(Some(35.4), 62.35, "office2pdf CLI Manual v0.6", arial(8.0));

    let baselines: Vec<f64> = baselines_of(&doc, "office2pdf CLI Manual");
    assert_eq!(baselines.len(), 1, "the header is one line");
    let expected_pt: f64 = 35.4 + ascender_em * 8.0;
    assert!(
        (baselines[0] - expected_pt).abs() < 0.01,
        "header baseline {}pt should be {expected_pt}pt",
        baselines[0]
    );
    assert!(
        !shaped_by(&doc, "office2pdf CLI Manual", &ARIAL_METRIC_FACES)
            || (baselines[0] - 42.72).abs() < 0.12,
        "Word's own export measures 42.72pt, not {}pt",
        baselines[0]
    );

    // Word keeps the hhea line gap above the header origin, so the header
    // ascent is not the body line's gap-inclusive one.
    let (body_ascent_em, _, _) = crate::render::pdf::font_line_metrics_em("Arial")
        .expect("Arial line metrics should resolve");
    assert!(
        (baselines[0] - (35.4 + body_ascent_em * 8.0)).abs() > 0.2,
        "the body line's ascent would put the header baseline at {}pt",
        35.4 + body_ascent_em * 8.0
    );
}

#[cfg(not(target_arch = "wasm32"))]
/// The same placement at a different header distance and font size: the ascent
/// scales with the size and the origin follows `w:header`, so neither term can
/// be a constant.
#[test]
fn test_header_first_baseline_scales_with_font_size_and_header_distance() {
    let ascender_em: f64 =
        crate::render::pdf::font_hhea_ascender_em("Arial").expect("Arial metrics should resolve");
    let doc = doc_with_header_run(Some(56.7), 85.05, "Datasheet", arial(12.0));

    let baselines: Vec<f64> = baselines_of(&doc, "Datasheet");
    let expected_pt: f64 = 56.7 + ascender_em * 12.0;
    assert_eq!(baselines.len(), 1, "the header is one line");
    assert!(
        (baselines[0] - expected_pt).abs() < 0.01,
        "header baseline {}pt should be {expected_pt}pt",
        baselines[0]
    );
}

#[cfg(not(target_arch = "wasm32"))]
/// A header line carrying East Asian text keeps the extra ascent Word gives it:
/// half of the 30% its line gains over the font's own (issues #518, #629).
/// `03_meeting_minutes_ko` and `10_research_report_ko` both measure 45.60pt at
/// the same 35.40pt distance where an 8pt Arial header measures 42.72pt.
#[test]
fn test_east_asian_header_baseline_keeps_the_word_line_bonus() {
    let Some(ascender_em) = crate::render::pdf::font_hhea_ascender_em("Malgun Gothic") else {
        return;
    };
    let (_, _, pitch_em) = crate::render::pdf::font_line_metrics_em("Malgun Gothic")
        .expect("a resolved face has line metrics");
    let doc = doc_with_header_run(
        Some(35.4),
        62.35,
        "회의록 | 사내 문서",
        TextStyle {
            font_family: Some("Malgun Gothic".to_string()),
            east_asian_font_family: Some("Malgun Gothic".to_string()),
            font_size: Some(8.0),
            ..TextStyle::default()
        },
    );

    let baselines: Vec<f64> = baselines_of(&doc, "회의록");
    let expected_pt: f64 = 35.4 + (ascender_em + 0.15 * pitch_em) * 8.0;
    assert_eq!(baselines.len(), 1, "the header is one line");
    assert!(
        (baselines[0] - expected_pt).abs() < 0.01,
        "East Asian header baseline {}pt should be {expected_pt}pt",
        baselines[0]
    );
    assert!(
        (baselines[0] - (35.4 + ascender_em * 8.0)).abs() > 1.0,
        "the bonus must lift the baseline clear of the bare ascent"
    );
    assert!(
        !shaped_by(&doc, "회의록", &MALGUN_METRIC_FACES) || (baselines[0] - 45.60).abs() < 0.15,
        "Word's own export measures 45.60pt, not {}pt",
        baselines[0]
    );
}

#[cfg(not(target_arch = "wasm32"))]
/// The face decides the header ascent, not the script of the line's
/// characters: a Latin-only header set in Malgun Gothic keeps the same East
/// Asian bonus a Korean one gets — the rule the body line took in issue #643
/// and the footer in issue #630 (issue #814).
///
/// Measured on a native export: `10_research_report_ko` with its header text
/// replaced by `Monthly Customer Satisfaction Trend Report` — the only patched
/// factor — keeps its first baseline at 45.60pt at `w:header="708"` = 35.40pt,
/// exactly where the Korean control's sits, where the bare hhea ascender would
/// seat it at 44.11pt.
#[test]
fn test_latin_only_header_in_east_asian_face_keeps_the_word_line_bonus() {
    let Some(ascender_em) = crate::render::pdf::font_hhea_ascender_em("Malgun Gothic") else {
        return;
    };
    let (_, _, pitch_em) = crate::render::pdf::font_line_metrics_em("Malgun Gothic")
        .expect("a resolved face has line metrics");
    let doc = doc_with_header_run(
        Some(35.4),
        62.35,
        "Monthly Customer Satisfaction Trend Report",
        TextStyle {
            font_family: Some("Malgun Gothic".to_string()),
            east_asian_font_family: Some("Malgun Gothic".to_string()),
            font_size: Some(8.0),
            ..TextStyle::default()
        },
    );

    let baselines: Vec<f64> = baselines_of(&doc, "Monthly Customer Satisfaction");
    let expected_pt: f64 = 35.4 + (ascender_em + 0.15 * pitch_em) * 8.0;
    assert_eq!(baselines.len(), 1, "the header is one line");
    assert!(
        (baselines[0] - expected_pt).abs() < 0.01,
        "Latin header in a CJK face: baseline {}pt should be {expected_pt}pt",
        baselines[0]
    );
    assert!(
        (baselines[0] - (35.4 + ascender_em * 8.0)).abs() > 1.0,
        "the bonus must lift the baseline clear of the bare ascent"
    );
    assert!(
        !shaped_by(&doc, "Monthly Customer Satisfaction", &MALGUN_METRIC_FACES)
            || (baselines[0] - 45.60).abs() < 0.15,
        "Word's own export measures 45.60pt, not {}pt",
        baselines[0]
    );
}

#[cfg(not(target_arch = "wasm32"))]
/// Moving the band must not touch the story's own line advance.
///
/// The header ascent is a property of where the *first* line sits, not of how
/// far apart the lines are. Declaring it as a `top-edge` on the paragraph made
/// Typst widen every wrapped line's box and stretched an 8pt Arial header's
/// advance from 10.93pt to 12.44pt; shifting the band leaves every box alone
/// (issue #629).
#[test]
fn test_shifting_the_header_band_leaves_the_wrapped_line_advance_alone() {
    let ascender_em: f64 =
        crate::render::pdf::font_hhea_ascender_em("Arial").expect("Arial metrics should resolve");
    // Every line has to carry the marker so both wrapped lines are found.
    let wrapping: String = "office2pdf ".repeat(20);

    let pinned: Vec<f64> = baselines_of(
        &doc_with_header_run(Some(35.4), 62.35, &wrapping, arial(8.0)),
        "office2pdf",
    );
    let unpinned: Vec<f64> = baselines_of(
        &doc_with_header_run(None, 62.35, &wrapping, arial(8.0)),
        "office2pdf",
    );

    assert_eq!(
        pinned.len(),
        2,
        "the header paragraph must wrap: {pinned:?}"
    );
    assert_eq!(unpinned.len(), 2, "the header paragraph must wrap");
    let pinned_advance: f64 = pinned[1] - pinned[0];
    let unpinned_advance: f64 = unpinned[1] - unpinned[0];
    assert!(
        (pinned_advance - unpinned_advance).abs() < 0.001,
        "the band shift changed the wrapped advance: {pinned_advance} vs {unpinned_advance}"
    );
    assert!(
        (pinned[0] - (35.4 + ascender_em * 8.0)).abs() < 0.01,
        "the first line still has to land on Word's baseline, not {}pt",
        pinned[0]
    );
    assert!(
        pinned[0] > unpinned[0] + 1.0,
        "the pinned header must actually have moved"
    );
}

#[cfg(not(target_arch = "wasm32"))]
/// The band is sized by the first paragraph the story *emits*, whatever it is
/// made of.
///
/// A `PAGE` field carries its own run properties, so a header whose first
/// paragraph is nothing but a page number still has an ascent to seat. Reading
/// only text runs left the decision unconsumed and handed it to the second
/// paragraph, which then seated the wrong line (issue #629).
#[test]
fn test_header_whose_first_paragraph_is_a_page_field_seats_that_line() {
    use crate::ir::{HFInline, HeaderFooterParagraph};

    let ascender_em: f64 =
        crate::render::pdf::font_hhea_ascender_em("Arial").expect("Arial metrics should resolve");
    let page_field = HeaderFooterParagraph {
        style: ParagraphStyle::default(),
        elements: vec![HFInline::PageNumber(arial(8.0))],
        border: None,
        border_space: None,
        sheet_section_is_rich: false,
        frame: None,
    };
    let second = header_text_paragraph("office2pdf CLI Manual v0.6", arial(8.0));

    let pinned = doc_with_header(Some(35.4), 62.35, vec![page_field.clone(), second.clone()]);
    let unpinned = doc_with_header(None, 62.35, vec![page_field, second]);

    let pinned_number: Vec<f64> = baselines_of(&pinned, "1");
    let expected_pt: f64 = 35.4 + ascender_em * 8.0;
    assert!(
        pinned_number
            .first()
            .is_some_and(|first| (first - expected_pt).abs() < 0.01),
        "the page-number line should sit at {expected_pt}pt, not {pinned_number:?}"
    );

    // The second paragraph rides along; the gap between the two is the story's
    // own and must survive the shift.
    let pinned_second: f64 = baselines_of(&pinned, "office2pdf CLI Manual")[0];
    let unpinned_number: f64 = baselines_of(&unpinned, "1")[0];
    let unpinned_second: f64 = baselines_of(&unpinned, "office2pdf CLI Manual")[0];
    assert!(
        ((pinned_second - pinned_number[0]) - (unpinned_second - unpinned_number)).abs() < 0.001,
        "the shift changed the story's paragraph advance"
    );
}

#[cfg(not(target_arch = "wasm32"))]
/// Header ink never reaches the body's first line.
///
/// This once asserted the ink stayed inside the *declared* `w:top - w:header`
/// band, because the shift was clamped to whatever slack the story left — a
/// stand-in noted as holding only "until #736 is modelled". #736 is modelled
/// now: the band grows with the story instead, so the invariant worth keeping
/// is the one the clamp was protecting, that the header never overprints the
/// body.
#[test]
fn test_header_ink_never_reaches_the_body() {
    if crate::render::pdf::font_hhea_ascender_em("Malgun Gothic").is_none() {
        return;
    }
    let korean = |text: &str| {
        header_text_paragraph(
            text,
            TextStyle {
                font_family: Some("Malgun Gothic".to_string()),
                east_asian_font_family: Some("Malgun Gothic".to_string()),
                font_size: Some(12.0),
                ..TextStyle::default()
            },
        )
    };
    let doc = doc_with_header(
        Some(35.4),
        62.35,
        vec![
            korean("주식회사 오피스투피디에프 기술연구소"),
            korean("서울특별시 강남구 테헤란로 000, 00층"),
        ],
    );

    let last_header_baseline: f64 = *baselines_of(&doc, "서울특별시")
        .last()
        .expect("the second header line is placed");
    let first_body_baseline: f64 = *baselines_of(&doc, "Body")
        .first()
        .expect("the body's first line is placed");
    assert!(
        last_header_baseline < first_body_baseline,
        "header ink reached {last_header_baseline}pt, at or past the body's \
         first baseline at {first_body_baseline}pt"
    );
    let body_baseline: f64 = baselines_of(&doc, "Body")[0];
    assert!(
        last_header_baseline < body_baseline,
        "the header overprints the body's first line at {body_baseline}pt"
    );
}

#[cfg(not(target_arch = "wasm32"))]
/// The footer keeps its descender anchor: `w:pgMar/@w:footer` measures to the
/// bottom of the footer, so no ascent is involved (issue #630 tracks its own
/// placement). Adding a shifted header must not disturb it.
#[test]
fn test_header_band_shift_leaves_the_footer_where_it_was() {
    use crate::ir::{HeaderFooter, HeaderFooterParagraph};

    let footer_paragraph: HeaderFooterParagraph = header_text_paragraph("- 1 -", arial(8.0));
    let page = |header: Option<HeaderFooter>| {
        make_doc(vec![Page::Flow(FlowPage {
            first_header: None,
            first_footer: None,
            size: PageSize::default(),
            margins: Margins {
                top: 62.35,
                bottom: 62.35,
                left: 70.85,
                right: 70.85,
            },
            content: vec![make_paragraph("Body")],
            header,
            footer: Some(HeaderFooter {
                shapes: Vec::new(),
                distance_from_edge: Some(35.4),
                sheet_print_scale: None,
                paragraphs: vec![footer_paragraph.clone()],
            }),
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })])
    };

    let without_header = page(None);
    let with_header = page(Some(HeaderFooter {
        shapes: Vec::new(),
        distance_from_edge: Some(35.4),
        sheet_print_scale: None,
        paragraphs: vec![header_text_paragraph("Header", arial(8.0))],
    }));

    let alone: f64 = baselines_of(&without_header, "- 1 -")[0];
    let alongside: f64 = baselines_of(&with_header, "- 1 -")[0];
    assert!(
        (alone - alongside).abs() < 0.001,
        "the header shift moved the footer from {alone}pt to {alongside}pt"
    );
    let output = generate_typst(&with_header).unwrap();
    assert!(
        output.source.contains("footer-descent: 0pt"),
        "the footer origin must stay on the bottom margin line"
    );
}

#[cfg(not(target_arch = "wasm32"))]
/// A header without `w:pgMar/@w:header` has no origin to measure an ascent
/// from, so its line stays where the renderer seats it.
#[test]
fn test_unpinned_header_keeps_the_renderer_seat() {
    let ascender_em: f64 =
        crate::render::pdf::font_hhea_ascender_em("Arial").expect("Arial metrics should resolve");
    let cap_height_em: f64 =
        crate::render::pdf::font_cap_height_em("Arial").expect("Arial metrics should resolve");
    assert!(
        (ascender_em - cap_height_em).abs() > 0.05,
        "the two seats must differ for this test to mean anything"
    );

    let doc = doc_with_header_run(None, 62.35, "Header", arial(8.0));
    let output = generate_typst(&doc).expect("document should generate");
    assert!(
        !output.source.contains("place(top, dy:"),
        "an unpinned header must not be shifted: {}",
        output.source
    );
    let baseline: f64 = baselines_of(&doc, "Header")[0];
    assert!(
        baseline < 62.35,
        "the unpinned header still sits above the top margin, not at {baseline}pt"
    );
}

/// Word applies the containing run's properties to a `PAGE` field result, so
/// the number must render in the run's font, size, and color rather than the
/// document default.
#[test]
fn test_page_number_field_uses_its_run_style() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};

    let field_style = TextStyle {
        font_size: Some(8.0),
        color: Some(Color::new(0x88, 0x88, 0x88)),
        ..TextStyle::default()
    };
    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![
                    HFInline::Run(Run {
                        text: "- ".to_string(),
                        style: field_style.clone(),
                        href: None,
                        footnote: None,
                    }),
                    HFInline::PageNumber(field_style.clone()),
                ],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    let counter = output
        .source
        .find(r#"#context counter(page).display("1")"#)
        .expect("page counter emitted");
    let prefix = &output.source[..counter];
    let wrapper = prefix
        .rfind("#text(")
        .expect("the counter is wrapped in its run's text properties");
    assert!(
        prefix[wrapper..].contains("size: 8pt"),
        "the field keeps the run's size, got: {}",
        &prefix[wrapper..]
    );
    assert!(
        prefix[wrapper..].contains("rgb(136, 136, 136)"),
        "the field keeps the run's color"
    );
}

/// An unstyled field stays a bare counter, so documents that never style their
/// page numbers are unchanged.
#[test]
fn test_unstyled_page_number_field_stays_bare() {
    use crate::ir::{HFInline, HeaderFooter, HeaderFooterParagraph};

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::PageNumber(TextStyle::default())],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            }],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains(r#"#context counter(page).display("1")"#)
    );
    assert!(
        !output.source.contains("#text()[#counter"),
        "no empty text wrapper"
    );
}

#[test]
fn test_section_page_numbering_updates_the_counter_and_its_numerals() {
    // Word restarts the counter at the section boundary and renders the
    // numerals w:fmt names; Typst counts from the document start in decimal
    // unless told otherwise (issue #582).
    let Page::Flow(mut flow) = make_flow_page(vec![make_paragraph("front matter")]) else {
        unreachable!()
    };
    flow.footer = Some(crate::ir::HeaderFooter {
        shapes: Vec::new(),
        paragraphs: vec![crate::ir::HeaderFooterParagraph {
            style: ParagraphStyle::default(),
            elements: vec![HFInline::PageNumber(TextStyle::default())],
            border: None,
            border_space: None,
            sheet_section_is_rich: false,
            frame: None,
        }],
        distance_from_edge: None,
        sheet_print_scale: None,
    });
    flow.page_numbering = Some(crate::ir::PageNumbering {
        start: Some(1),
        format: crate::ir::PageNumberFormat::LowerRoman,
    });

    let output = generate_typst(&make_doc(vec![Page::Flow(flow)])).unwrap();

    assert!(
        output.source.contains("#counter(page).update(1)"),
        "the section restarts the counter: {}",
        output.source
    );
    assert!(
        output
            .source
            .contains(r#"#context counter(page).display("i")"#),
        "the PAGE field renders the section's numerals: {}",
        output.source
    );
}

#[test]
fn test_contents_block_emits_an_outline_at_its_declared_depth() {
    // The entries, their page numbers, and the leaders between them all come
    // from where the headings land, which only the layout knows (issue #576).
    let doc = make_doc(vec![make_flow_page(vec![
        Block::TableOfContents(crate::ir::TableOfContents::Headings { depth: 3 }),
        Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                heading_level: Some(1),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "1. 개요".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        }),
    ])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("query(<o2p-toc>)") && output.source.contains("level <= 3"),
        "the contents block resolves against the document's headings, to its          declared depth: {}",
        output.source
    );
    assert!(
        output
            .source
            .contains("#metadata((level: 1, text: \"1. 개요\""),
        "each heading drops the plain text its entry is built from: {}",
        output.source
    );
}

#[test]
fn test_caption_list_queries_the_captions_it_collects() {
    // A caption is not a heading, so Typst's outline cannot reach it. Each one
    // drops an invisible marker as it is laid out and the list queries those,
    // so both the entries and their page numbers come from the layout
    // (issue #576).
    let doc = make_doc(vec![make_flow_page(vec![
        Block::TableOfContents(crate::ir::TableOfContents::Captions {
            identifier: "Figure".to_string(),
        }),
        Block::Caption(crate::ir::Caption {
            identifier: "Figure".to_string(),
            entry_text: "변환 파이프라인".to_string(),
            paragraph: Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "그림 1  변환 파이프라인".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            },
        }),
    ])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output
            .source
            .contains("#metadata[변환 파이프라인]<o2p-seq-Figure>"),
        "the caption carries the marker its list queries: {}",
        output.source
    );
    assert!(
        output.source.contains("query(<o2p-seq-Figure>)")
            && output.source.contains("let target = entry.location()")
            && output.source.contains("counter(page).at(target)"),
        "the list resolves each entry's page from where it landed: {}",
        output.source
    );
    // The caption's own runs still render; the auto-space marker sits between
    // its number and the Korean that follows, which is why this looks for the
    // two halves rather than the joined string. Each Korean eojeol carries the
    // frame that keeps it whole across a line break (issue #626), so the
    // rendered halves are matched in that form while the list entry, built
    // from `#metadata`, keeps the plain string.
    assert!(
        output.source.contains("#box[그림] 1")
            && output.source.contains("#box[변환] #box[파이프라인]")
            && output.source.contains("#metadata[변환 파이프라인]"),
        "the caption still renders as itself, beside its list entry: {}",
        output.source
    );
}

#[test]
fn test_a_caption_identifier_outside_ascii_still_labels() {
    // A `SEQ` name may be written in any script; a Typst label may not.
    let doc = make_doc(vec![make_flow_page(vec![Block::TableOfContents(
        crate::ir::TableOfContents::Captions {
            identifier: "표".to_string(),
        },
    )])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("query(<o2p-seq--d45c>)"),
        "the identifier reduces to label characters: {}",
        output.source
    );
}

/// Word numbers a contents entry the way the section it points into numbers
/// its pages, so an entry landing in roman-numbered front matter reads `i`
/// rather than `1` (issue #605). The format travels with the layout, because
/// only the layout knows which section an entry resolved into.
#[test]
fn test_contents_entries_number_in_the_target_sections_format() {
    use crate::ir::{PageNumberFormat, PageNumbering};

    let front_matter = Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![
            Block::TableOfContents(crate::ir::TableOfContents::Headings { depth: 3 }),
            Block::Paragraph(Paragraph {
                style: ParagraphStyle {
                    heading_level: Some(1),
                    ..ParagraphStyle::default()
                },
                runs: vec![Run {
                    text: "적용 범위".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            }),
        ],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: Some(PageNumbering {
            start: Some(1),
            format: PageNumberFormat::LowerRoman,
        }),
    });
    let body = Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                heading_level: Some(1),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "1. 개요".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: Some(PageNumbering {
            start: Some(1),
            format: PageNumberFormat::Decimal,
        }),
    });

    let output = generate_typst(&make_doc(vec![front_matter, body])).unwrap();

    assert!(
        output.source.contains("state(\"o2p-page-format\""),
        "the section's numeral format is recorded where the layout can read it back: {}",
        output.source
    );
    assert!(
        output.source.contains("o2p-page-format.update(\"i\")"),
        "the roman front matter records its format: {}",
        output.source
    );
    assert!(
        output.source.contains("o2p-page-format.update(\"1\")"),
        "the decimal body records its own: {}",
        output.source
    );
    assert!(
        output.source.contains("show outline.entry"),
        "the outline renders each entry's number through the recorded format: {}",
        output.source
    );
}

/// A caption list is numbered the same way a heading outline is: an entry
/// pointing at a table in roman-numbered front matter reads `i`, not `1`
/// (issue #605). The list builds its own rows, so it has to read the format
/// back at the entry's location just as the outline rule does.
#[test]
fn test_caption_list_numbers_in_the_target_sections_format() {
    use crate::ir::{Caption, PageNumberFormat, PageNumbering};

    let page = Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![
            Block::TableOfContents(crate::ir::TableOfContents::Captions {
                identifier: "표".to_string(),
            }),
            Block::Caption(Caption {
                identifier: "표".to_string(),
                entry_text: "문서 서지 정보".to_string(),
                paragraph: Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "표 1 문서 서지 정보".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                },
            }),
        ],
        header: None,
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: Some(PageNumbering {
            start: Some(1),
            format: PageNumberFormat::LowerRoman,
        }),
    });

    let output = generate_typst(&make_doc(vec![page])).unwrap();

    assert!(
        output.source.contains("o2p-page-format.at(target)"),
        "the caption list reads the format back at each entry's location: {}",
        output.source
    );
    assert!(
        !output.source.contains("#entry_page]"),
        "the raw page count no longer reaches the row: {}",
        output.source
    );
}

/// A footer whose run names an East Asian face, otherwise identical to
/// [`arial`]'s.
#[cfg(not(target_arch = "wasm32"))]
fn malgun(size_pt: f64) -> TextStyle {
    TextStyle {
        font_family: Some("Malgun Gothic".to_string()),
        font_size: Some(size_pt),
        ..TextStyle::default()
    }
}

/// Build a one-section document whose footer holds a single run, at
/// `w:pgMar/@w:footer` = 35.40pt on A4.
#[cfg(not(target_arch = "wasm32"))]
fn doc_with_footer_run(text: &str, style: TextStyle) -> Document {
    doc_with_spaced_footer_run(text, style, None)
}

/// The same footer, with the paragraph's resolved `w:spacing w:after` stated.
#[cfg(not(target_arch = "wasm32"))]
fn doc_with_spaced_footer_run(
    text: &str,
    style: TextStyle,
    space_after_pt: Option<f64>,
) -> Document {
    use crate::ir::HeaderFooter;

    let mut paragraph = header_text_paragraph(text, style);
    paragraph.style.space_after = space_after_pt;

    make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins {
            top: 62.35,
            bottom: 62.35,
            left: 70.85,
            right: 70.85,
        },
        content: vec![make_paragraph("Body")],
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: Some(35.4),
            sheet_print_scale: None,
            paragraphs: vec![paragraph],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })])
}

#[cfg(not(target_arch = "wasm32"))]
/// Word reserves the footer's last paragraph's own `w:spacing w:after` between
/// its last line and the `w:footer` anchor, so a stated gap lifts the whole
/// band by exactly that much.
///
/// `tests/fixtures/docx/unit_test_headers.docx` states no `w:pPrDefault`, so
/// its footer resolves Word's built-in `Normal` `w:after="160"` = 8pt; the
/// native export puts that footer baseline 46.56pt above the page bottom
/// against the 38.54pt an unreserved band produces — the gap is that 8pt
/// (issue #1195).
#[test]
fn test_footer_band_reserves_the_last_paragraph_space_after() {
    let unreserved: f64 = baselines_of(
        &doc_with_spaced_footer_run("- 1 -", arial(8.0), Some(0.0)),
        "- 1 -",
    )[0];

    for reserved_pt in [8.0, 16.0] {
        let baseline: f64 = baselines_of(
            &doc_with_spaced_footer_run("- 1 -", arial(8.0), Some(reserved_pt)),
            "- 1 -",
        )[0];
        assert!(
            (unreserved - baseline - reserved_pt).abs() < 0.01,
            "a {reserved_pt}pt `w:after` must lift the footer by {reserved_pt}pt: \
             {unreserved}pt unreserved against {baseline}pt reserved"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
/// A story stating no gap at all — every non-DOCX footer, whose paragraphs
/// carry no `w:spacing` — keeps the band it always had.
#[test]
fn test_footer_band_without_a_stated_space_after_is_unchanged() {
    let unstated: f64 = baselines_of(
        &doc_with_spaced_footer_run("- 1 -", arial(8.0), None),
        "- 1 -",
    )[0];
    let zero: f64 = baselines_of(
        &doc_with_spaced_footer_run("- 1 -", arial(8.0), Some(0.0)),
        "- 1 -",
    )[0];

    assert!(
        (unstated - zero).abs() < 0.01,
        "an unstated gap must seat the band where a zero one does: \
         {unstated}pt against {zero}pt"
    );
}

#[cfg(not(target_arch = "wasm32"))]
/// The footer's last baseline is one line-box descent above the `w:footer`
/// edge, and that descent is the resolved face's — not a constant.
///
/// The three golden mocks 01, 02 and 03 differ in their `footer1.xml` only in
/// `w:rFonts`, and Word moves the footer baseline with the font: 804.72pt for
/// Arial against 802.80pt for Malgun Gothic. Typst's `bottom-edge:
/// "descender"` is nearly right for Arial and 2.10pt wrong for Malgun Gothic,
/// so a test that pinned only one font would have passed throughout
/// (issue #630).
///
/// Needs a Korean face to say anything: where none is installed — the Linux CI
/// runner is one — the East Asian side has no metrics, the band keeps the
/// renderer's own seat, and there is no font-driven difference to measure. The
/// emission itself is covered unconditionally by
/// [`test_footer_band_states_its_own_bottom_edge`].
#[test]
fn test_footer_baseline_follows_its_own_font_descent() {
    if crate::render::pdf::font_line_metrics_em("Malgun Gothic").is_none() {
        return;
    }

    let latin: f64 = baselines_of(&doc_with_footer_run("- 1 -", arial(8.0)), "- 1 -")[0];
    let east_asian: f64 = baselines_of(&doc_with_footer_run("- 1 -", malgun(8.0)), "- 1 -")[0];

    assert!(
        east_asian < latin,
        "an East Asian footer carries more below its baseline, so it must sit \
         higher than the Arial one: Malgun {east_asian}pt against Arial {latin}pt"
    );
    // Word's own gap between the two, from the exports named above.
    let gap: f64 = latin - east_asian;
    assert!(
        (gap - 1.92).abs() < 0.35,
        "Word separates the two footers by 1.92pt; this build separates them by {gap}pt"
    );
}

#[cfg(not(target_arch = "wasm32"))]
/// Triangulation for the emission: the band states the descent itself rather
/// than deferring to the renderer's `"descender"`, which is the *normalised*
/// one and so answers a different question.
///
/// Arial, because every runner resolves it — through Liberation Sans where the
/// face itself is absent — so this half of #630 is pinned everywhere.
#[test]
fn test_footer_band_states_its_own_bottom_edge() {
    let (ascender_em, descender_em, pitch_em) =
        crate::render::pdf::font_line_metrics_em("Arial").expect("Arial metrics should resolve");
    let expected_em: f64 = pitch_em - ascender_em;
    assert!(
        (expected_em - descender_em).abs() < 1e-9,
        "a Latin line's sub-baseline share is its descender; the model changed"
    );

    let source = generate_typst(&doc_with_footer_run("- 1 -", arial(8.0)))
        .expect("document should generate")
        .source;

    assert!(
        !source.contains("bottom-edge: \"descender\""),
        "the footer band must not take the renderer's normalised descender: {source}"
    );
    assert!(
        source.contains(&format!("bottom-edge: -{}em", format_f64(expected_em))),
        "the band must state the face's own {expected_em}em descent: {source}"
    );
    assert!(
        source.contains("footer-descent: 0pt"),
        "the footer origin must stay on the bottom margin line: {source}"
    );
}

/// A header rule is spaced from the line's bottom, not the font's descender.
///
/// Regression for #737: Typst's `"descender"` is its *normalised* descender,
/// 0.199em for Malgun Gothic against the 0.4412em its 1.3x line box actually
/// carries, so a Korean header's rule sat 1.98pt high. The assertion is font
/// independent on purpose — CI's Linux runner has no CJK face, so it checks
/// that the header asks [`word_line_box_descent_em`] rather than that the
/// answer is any particular number.
#[test]
fn a_header_rule_is_spaced_from_the_line_box_bottom() {
    use crate::ir::{BorderSide, CellBorder, HFInline, HeaderFooter, HeaderFooterParagraph};

    let run = Run {
        text: "Minutes".to_string(),
        style: TextStyle::default(),
        href: None,
        footnote: None,
    };
    let paragraph = HeaderFooterParagraph {
        style: ParagraphStyle::default(),
        elements: vec![HFInline::Run(run.clone())],
        border: Some(CellBorder {
            top: None,
            bottom: Some(BorderSide {
                width: 0.5,
                color: Color::new(0xCC, 0xCC, 0xCC),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
            }),
            left: None,
            right: None,
        }),
        border_space: None,
        sheet_section_is_rich: false,
        frame: None,
    };
    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![paragraph],
        }),
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let source = generate_typst(&doc).unwrap().source;
    let expected: String = crate::render::typst_gen::text::word_line_box_descent_em(&[run])
        .map(|descent_em| format!("bottom-edge: -{}em", format_f64(descent_em)))
        .unwrap_or_else(|| "bottom-edge: \"descender\"".to_string());
    assert!(
        source.contains(&expected),
        "the header rule must be spaced from the line box bottom ({expected}): {source}"
    );
}

/// A header taller than its band grows the top margin (issue #736).
///
/// The two-line case above was only clamped, so it never reached the body and
/// would pass without this fix. Four 12pt lines into the same 26.95pt band do
/// overflow: before the margin grew, the third and fourth header lines
/// interleaved with the body text, which the reference export places below all
/// four.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_header_taller_than_its_band_pushes_the_body_down() {
    if crate::render::pdf::font_hhea_ascender_em("Malgun Gothic").is_none() {
        return;
    }
    let line = |text: &str| {
        header_text_paragraph(
            text,
            TextStyle {
                font_family: Some("Malgun Gothic".to_string()),
                east_asian_font_family: Some("Malgun Gothic".to_string()),
                font_size: Some(12.0),
                ..TextStyle::default()
            },
        )
    };
    // 26.95pt of band against four 12pt East Asian lines.
    let doc = doc_with_header(
        Some(35.4),
        62.35,
        vec![
            line("첫째 줄"),
            line("둘째 줄"),
            line("셋째 줄"),
            line("넷째 줄"),
        ],
    );

    let last_header_baseline: f64 = *baselines_of(&doc, "넷째 줄")
        .last()
        .expect("the fourth header line is placed");
    let first_body_baseline: f64 = *baselines_of(&doc, "Body")
        .first()
        .expect("the body's first line is placed");
    assert!(
        last_header_baseline < first_body_baseline,
        "the fourth header line sits at {last_header_baseline}pt, at or past \
         the body's first baseline at {first_body_baseline}pt — the top margin \
         did not grow"
    );
    // And the growth is real rather than the body merely starting late.
    assert!(
        first_body_baseline > 62.35,
        "the body must be pushed past the declared 62.35pt top margin, got \
         {first_body_baseline}pt"
    );
}

/// A taller first-page story grows the shared margin too (issues #736, #846).
///
/// One margin serves the whole section, so measuring only the default story
/// would leave a taller `w:titlePg` header overprinting page one — the same
/// defect, one page in.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_taller_first_page_header_also_grows_the_margin() {
    use crate::ir::HeaderFooter;

    if crate::render::pdf::font_hhea_ascender_em("Malgun Gothic").is_none() {
        return;
    }
    let line = |text: &str| {
        header_text_paragraph(
            text,
            TextStyle {
                font_family: Some("Malgun Gothic".to_string()),
                east_asian_font_family: Some("Malgun Gothic".to_string()),
                font_size: Some(12.0),
                ..TextStyle::default()
            },
        )
    };
    // The default story fits its band; only the first-page one overflows.
    let mut doc = doc_with_header(Some(35.4), 62.35, vec![line("한 줄")]);
    let Some(Page::Flow(page)) = doc.pages.first_mut() else {
        panic!("the fixture is a flow page");
    };
    page.first_header = Some(HeaderFooter {
        shapes: Vec::new(),
        distance_from_edge: Some(35.4),
        sheet_print_scale: None,
        paragraphs: vec![
            line("표지 첫째 줄"),
            line("표지 둘째 줄"),
            line("표지 셋째 줄"),
            line("표지 넷째 줄"),
        ],
    });

    let last_first_page_baseline: f64 = *baselines_of(&doc, "표지 넷째 줄")
        .last()
        .expect("the fourth first-page header line is placed");
    let first_body_baseline: f64 = *baselines_of(&doc, "Body")
        .first()
        .expect("the body's first line is placed");
    assert!(
        last_first_page_baseline < first_body_baseline,
        "the first-page header reaches {last_first_page_baseline}pt against the \
         body's {first_body_baseline}pt — the shared margin ignored it"
    );
}

/// A header line advances by Word's pitch, not Typst's default leading.
///
/// Regression for #735. The story carried no leading, so its lines took the
/// 0.65em default on top of Typst's cap-height edge — 10.9305pt for 8pt Arial
/// against Word's 9.1992pt. The leading is stated once for the story, which is
/// a single Typst paragraph joined by `\\` line breaks; stating it per
/// paragraph would make each one a block and Typst would put `par(spacing:)`
/// between them instead.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_header_line_advances_by_words_pitch() {
    let runs = vec![Run {
        text: "Header".to_string(),
        style: TextStyle {
            font_family: Some("Arial".to_string()),
            font_size: Some(8.0),
            ..TextStyle::default()
        },
        href: None,
        footnote: None,
    }];
    let Some(expected) = crate::render::typst_gen::text::word_hf_line_leading_pt(&runs, 0.0) else {
        return; // the face is unavailable on this runner
    };
    let doc = doc_with_header(
        Some(35.4),
        62.35,
        vec![header_text_paragraph("Header", runs[0].style.clone())],
    );
    let source = generate_typst(&doc).unwrap().source;

    let marker: String = format!("#set par(leading: {}pt)", format_f64(expected));
    assert_eq!(
        source.matches(marker.as_str()).count(),
        1,
        "the story states its leading exactly once, expected {marker}: {source}"
    );
    // Word's advance is the leading plus the cap-height edge it is measured
    // against, so the emitted value must be strictly less than the advance.
    assert!(
        expected > 0.0 && expected < 9.1992,
        "8pt Arial leading tops the cap-height edge up to Word's 9.1992pt \
         advance, got {expected}pt"
    );
}

/// A header story's banner carries no text, so it never reaches the paragraph
/// path — and `behindDoc="1"` puts it under the page's own content, which the
/// foreground layer cannot do (issue #961).
#[test]
fn a_behind_text_header_banner_is_drawn_on_the_background_layer() {
    use crate::ir::{
        FrameAnchor, GradientFill, GradientStop, HeaderFooter, HeaderFooterFrame,
        HeaderFooterShape, Shape, ShapeKind,
    };

    let banner = HeaderFooterShape {
        shape: Shape {
            kind: ShapeKind::Path {
                subpaths: vec![crate::ir::Subpath::closed_outline(vec![
                    (0.0, 0.0),
                    (1.0, 0.0),
                    (1.0, 0.65),
                    (0.0, 1.0),
                ])],
            },
            fill: None,
            gradient_fill: Some(GradientFill {
                stops: vec![
                    GradientStop {
                        position: 0.0,
                        color: Color::new(0x9F, 0xDF, 0xBF),
                    },
                    GradientStop {
                        position: 1.0,
                        color: Color::new(0x4E, 0xB3, 0xCF),
                    },
                ],
                angle: 32.0,
            }),
            pattern_fill: None,
            stroke: None,
            rotation_deg: None,
            opacity: None,
            shadow: None,
            top_bevel: None,
        },
        // Wider than the 595.28pt page, centred, so it hangs off both edges.
        width: 609.12,
        height: 327.6,
        frame: HeaderFooterFrame {
            wraps_text: true,
            x: None,
            y: None,
            width: Some(609.12),
            height: Some(327.6),
            horizontal_anchor: FrameAnchor::Page,
            vertical_anchor: FrameAnchor::Page,
            horizontal_align: Some(crate::ir::FrameAlign::Center),
            vertical_align: Some(crate::ir::FrameAlign::Start),
            inset_left: 0.0,
            inset_top: 0.0,
            bottom_offset: None,
        },
        behind_text: true,
    };

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: Some(HeaderFooter {
            shapes: vec![banner],
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: Vec::new(),
        }),
        footer: None,
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("background: ["),
        "under the body, not over it: {}",
        output.source
    );
    assert!(!output.source.contains("foreground: ["));
    // Centring a 609.12pt banner on a 595.28pt page overhangs by 6.92pt.
    assert!(
        output.source.contains("#place(top + left, dx: -6.92"),
        "{}",
        output.source
    );
    // The box gives `#rotate` a frame of the shape's own extent; without it a
    // turned banner is laid out against the page width instead.
    assert!(
        output
            .source
            .contains("[#box(width: 609.12pt, height: 327.6pt)[#curve(fill-rule: \"even-odd\""),
        "{}",
        output.source
    );
    assert!(
        output.source.contains(
            "gradient.linear((rgb(159, 223, 191), 0%), (rgb(78, 179, 207), 100%), angle: 32deg, space: rgb)"
        ),
        "{}",
        output.source
    );
}

/// A `#block` fills its region and wraps; a `#box` shrinks to its content and
/// does not. That is the whole difference between the one line `<a:bodyPr
/// wrap="none">` asks for and the two lines a 1.33pt overflow produced
/// (issue #967).
#[test]
fn a_non_wrapping_anchored_frame_sizes_to_its_content() {
    use crate::ir::{
        FrameAnchor, HFInline, HeaderFooter, HeaderFooterFrame, HeaderFooterParagraph,
    };

    let frame = |wraps_text: bool| HeaderFooterFrame {
        x: Some(20.0),
        y: Some(700.0),
        width: Some(65.8),
        height: None,
        horizontal_anchor: FrameAnchor::Page,
        vertical_anchor: FrameAnchor::Page,
        horizontal_align: None,
        vertical_align: None,
        inset_left: 0.0,
        inset_top: 0.0,
        bottom_offset: None,
        wraps_text,
    };
    let page = |wraps_text: bool| {
        Page::Flow(FlowPage {
            first_header: None,
            first_footer: None,
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![make_paragraph("Body")],
            header: None,
            footer: Some(HeaderFooter {
                shapes: Vec::new(),
                distance_from_edge: None,
                sheet_print_scale: None,
                paragraphs: vec![HeaderFooterParagraph {
                    style: ParagraphStyle::default(),
                    elements: vec![HFInline::Run(Run {
                        text: "Sensitivity: Internal".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    })],
                    border: None,
                    border_space: None,
                    sheet_section_is_rich: false,
                    frame: Some(frame(wraps_text)),
                }],
            }),
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })
    };

    let non_wrapping = generate_typst(&make_doc(vec![page(false)])).unwrap().source;
    assert!(non_wrapping.contains("[#box()["), "{non_wrapping}");
    // The column width must not reach the markup at all — stating it is what
    // makes Typst break the line.
    assert!(!non_wrapping.contains("65.8pt"), "{non_wrapping}");

    let wrapping = generate_typst(&make_doc(vec![page(true)])).unwrap().source;
    assert!(wrapping.contains("[#block(width: 65.8pt)["), "{wrapping}");
}

/// The #1370 reference's bottom-seated WPS text box pins the last line's em box
/// above its bottom inset; it does not put the text baseline directly on the
/// inset line.
///
/// `Place your event title here.docx` declares an 8pt one-line footer with a
/// 15pt `bIns`. Its reference PDF therefore puts the baseline about 23pt above
/// the Letter page bottom, while placing the baseline at only 15pt produces
/// the 8.05pt downward error tracked in issue #1370.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_bottom_seated_anchored_frame_keeps_one_em_above_its_bottom_inset() {
    use crate::ir::{
        FrameAlign, FrameAnchor, HFInline, HeaderFooter, HeaderFooterFrame, HeaderFooterParagraph,
    };

    for (declared_size, expected_size) in [(Some(8.0), 8.0), (Some(12.0), 12.0), (None, 11.0)] {
        let doc = make_doc(vec![Page::Flow(FlowPage {
            first_header: None,
            first_footer: None,
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![make_paragraph("Body")],
            header: None,
            footer: Some(HeaderFooter {
                shapes: Vec::new(),
                distance_from_edge: None,
                sheet_print_scale: None,
                paragraphs: vec![HeaderFooterParagraph {
                    style: ParagraphStyle::default(),
                    elements: vec![HFInline::Run(Run {
                        text: "Sensitivity: Internal".to_string(),
                        style: declared_size.map_or_else(TextStyle::default, arial),
                        href: None,
                        footnote: None,
                    })],
                    border: None,
                    border_space: None,
                    sheet_section_is_rich: false,
                    frame: Some(HeaderFooterFrame {
                        x: None,
                        y: None,
                        width: Some(65.8),
                        height: Some(25.55),
                        horizontal_anchor: FrameAnchor::Page,
                        vertical_anchor: FrameAnchor::Page,
                        horizontal_align: Some(FrameAlign::Start),
                        vertical_align: Some(FrameAlign::End),
                        inset_left: 20.0,
                        inset_top: 0.0,
                        bottom_offset: Some(15.0),
                        wraps_text: false,
                    }),
                }],
            }),
            columns: None,
            line_grid_pitch: None,
            line_grid_snaps_lines: false,
            page_numbering: None,
        })]);

        let baselines = baselines_of(&doc, "Sensitivity: Internal");
        assert_eq!(baselines.len(), 1, "the footer label is one line");
        let expected = PageSize::default().height - 15.0 - expected_size;
        assert!(
            (baselines[0] - expected).abs() < 0.01,
            "bottom-seated {expected_size}pt baseline {}pt should be {expected}pt",
            baselines[0]
        );
    }
}

/// The page-left-aligned WPS footer in the #1219 / PR #1407 reference declares
/// a 20pt left inset, but LibreOffice 26.2.5.2 seats the run origin at 20.15pt.
/// This is the frame's horizontal seat, independent of the already-matched
/// bottom baseline and natural text width (issue #1487).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_page_left_aligned_wps_footer_uses_the_writer_text_origin_seat() {
    use crate::ir::{
        FrameAlign, FrameAnchor, HFInline, HeaderFooter, HeaderFooterFrame, HeaderFooterParagraph,
    };

    let doc = make_doc(vec![Page::Flow(FlowPage {
        first_header: None,
        first_footer: None,
        size: PageSize::default(),
        margins: Margins::default(),
        content: vec![make_paragraph("Body")],
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge: None,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "Sensitivity: Internal".to_string(),
                    style: arial(8.0),
                    href: None,
                    footnote: None,
                })],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: Some(HeaderFooterFrame {
                    x: None,
                    y: None,
                    width: Some(65.8),
                    height: Some(25.55),
                    horizontal_anchor: FrameAnchor::Page,
                    vertical_anchor: FrameAnchor::Page,
                    horizontal_align: Some(FrameAlign::Start),
                    vertical_align: Some(FrameAlign::End),
                    inset_left: 20.0,
                    inset_top: 0.0,
                    bottom_offset: Some(15.0),
                    wraps_text: false,
                }),
            }],
        }),
        columns: None,
        line_grid_pitch: None,
        line_grid_snaps_lines: false,
        page_numbering: None,
    })]);

    let run = placed_runs(&doc)
        .into_iter()
        .find(|run| run.text.contains("Sensitivity: Internal"))
        .expect("the footer label should be laid out");
    assert!(
        (run.left_pt - 20.15).abs() < 0.01,
        "Writer seats the page-left-aligned footer at 20.15pt, got {}pt",
        run.left_pt
    );
    let expected_baseline_pt: f64 = PageSize::default().height - 15.0 - 8.0;
    assert!(
        (run.baseline_pt - expected_baseline_pt).abs() < 0.01,
        "the horizontal correction must not move the matched {expected_baseline_pt}pt baseline, got {}pt",
        run.baseline_pt
    );
}

/// A concrete `<wp:posOffset>` is already the requested page coordinate and
/// must not inherit the Writer-only seat used for a page-left alignment.
#[test]
fn an_explicit_header_footer_x_offset_does_not_take_the_writer_aligned_seat() {
    use crate::ir::{FrameAlign, FrameAnchor, HeaderFooterFrame};

    let frame = HeaderFooterFrame {
        x: Some(12.5),
        y: None,
        width: Some(65.8),
        height: Some(25.55),
        horizontal_anchor: FrameAnchor::Page,
        vertical_anchor: FrameAnchor::Page,
        horizontal_align: Some(FrameAlign::Start),
        vertical_align: None,
        inset_left: 20.0,
        inset_top: 0.0,
        bottom_offset: None,
        wraps_text: false,
    };

    assert_eq!(page_anchored_hf_text_origin_x(&frame, 612.0), 32.5);
}

/// A 5 × 60pt grid on A4 portrait with 0.7in margins: the probe workbook of
/// issue #1110, whose native Excel-for-Mac export puts the printed grid's
/// left edge at 146pt. Its 50.4pt sides reach the renderer on the whole point
/// Excel prints against (issue #1127).
fn centered_sheet_page(centers: bool, column_widths: Vec<f64>) -> Page {
    Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins {
            top: 54.0,
            bottom: 54.0,
            left: 50.0,
            right: 50.0,
        },
        table: Table {
            rows: vec![TableRow {
                minimum_height: None,
                cells: column_widths.iter().map(|_| TableCell::default()).collect(),
                height: None,
            }],
            column_widths,
            centers_between_print_margins: centers,
            ..Table::default()
        },
        header: None,
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    })
}

/// The inset the emitted `#pad` states, or `None` when the page emits none.
fn sheet_centering_inset_pt(source: &str) -> Option<f64> {
    let rest: &str = source.split_once("#pad(left: ")?.1;
    let value: &str = rest.split_once("pt)[")?.0;
    value.parse().ok()
}

#[test]
fn test_horizontally_centered_sheet_insets_the_grid_from_the_left_margin() {
    let source = generate_typst(&make_doc(vec![centered_sheet_page(true, vec![60.0; 5])]))
        .unwrap()
        .source;
    let inset_pt: f64 = sheet_centering_inset_pt(&source)
        .unwrap_or_else(|| panic!("a centred sheet must inset its grid: {source}"));

    // 595.28pt page, 50pt margins, 300pt grid: the exact centre is 147.64pt
    // from the page edge and Excel prints the grid at 146pt (issue #1110).
    let grid_left_pt: f64 = 50.0 + inset_pt;
    assert!(
        (grid_left_pt - 146.0).abs() < 1.0,
        "grid left edge {grid_left_pt}pt must land within 1pt of Excel's 146pt: {source}"
    );
}

#[test]
fn test_uncentered_sheet_keeps_its_grid_on_the_left_margin() {
    let source = generate_typst(&make_doc(vec![centered_sheet_page(false, vec![60.0; 5])]))
        .unwrap()
        .source;
    assert_eq!(
        sheet_centering_inset_pt(&source),
        None,
        "a sheet without printOptions horizontalCentered must print flush: {source}"
    );
}

#[test]
fn test_centered_sheet_wider_than_the_printable_width_is_not_inset() {
    // Nothing is left to centre once the grid fills the page, and a negative
    // inset would push the first column off the left margin.
    let source = generate_typst(&make_doc(vec![centered_sheet_page(true, vec![200.0; 5])]))
        .unwrap()
        .source;
    assert_eq!(
        sheet_centering_inset_pt(&source),
        None,
        "an overflowing grid must not be inset: {source}"
    );
}

/// The fit-to-page A3 sheet attached to issue #1538 has a 1,078.30pt printed
/// grid centred on a 1,190.55pt page with 50pt side and 54pt top margins. Excel
/// lays the paper box out in sheet space before applying the 0.82 print scale.
/// Its resulting paint origin is (-0.185, -0.70)pt from the converter's table
/// origin. The first painted boundary then lands at 55.125 - 0.185 + 8.20 =
/// 63.14pt, matching the pinned native trace rather than the old 63.325pt.
#[test]
fn a_fit_scaled_a3_sheet_uses_its_scaled_paper_space_body_seat() {
    let Page::Sheet(mut sheet) = centered_sheet_page(true, vec![1_078.3]) else {
        unreachable!("centered_sheet_page builds a sheet page")
    };
    sheet.size = PageSize {
        width: 1_190.55,
        height: 841.89,
    };
    sheet.table.print_scale = Some(0.82);

    let source = generate_typst(&make_doc(vec![Page::Sheet(sheet)]))
        .unwrap()
        .source;
    assert!(
        source.contains("#move(dx: -0.185pt, dy: -0.7pt)["),
        "the scaled paint must start from Excel's sheet-space paper origin: {source}"
    );
    assert!(
        source.contains("#move(dx: 0.185pt, dy: 0.7pt)["),
        "cell content must keep its independently matched seat: {source}"
    );
    let inset_pt = sheet_centering_inset_pt(&source)
        .unwrap_or_else(|| panic!("the centred sheet needs its physical inset: {source}"));
    assert!((inset_pt - 5.125).abs() < 1e-9, "got {inset_pt}pt");
}

/// One-factor control for issue #1538: the same paper, margins and grid without
/// a fit scale keep the existing paper-space origin and centring calculation.
#[test]
fn an_unscaled_a3_sheet_does_not_take_the_scaled_body_seat() {
    let Page::Sheet(mut sheet) = centered_sheet_page(true, vec![1_078.3]) else {
        unreachable!("centered_sheet_page builds a sheet page")
    };
    sheet.size = PageSize {
        width: 1_190.55,
        height: 841.89,
    };

    let source = generate_typst(&make_doc(vec![Page::Sheet(sheet)]))
        .unwrap()
        .source;
    assert!(
        !source.contains("#move(dx: -0.185pt, dy: -0.7pt)["),
        "an unscaled sheet must keep its existing body origin: {source}"
    );
    let inset_pt = sheet_centering_inset_pt(&source)
        .unwrap_or_else(|| panic!("the unscaled centred sheet needs its normal inset: {source}"));
    assert!((inset_pt - 5.125).abs() < 1e-9, "got {inset_pt}pt");
}

#[test]
fn test_centered_sheet_moves_its_drawings_with_the_grid() {
    // Excel centres the printed sheet whole: a drawing floating over the
    // cells keeps its position relative to them (issue #1110).
    let Page::Sheet(mut sheet) = centered_sheet_page(true, vec![60.0; 5]) else {
        unreachable!("centered_sheet_page builds a sheet page")
    };
    sheet.text_boxes.push(crate::ir::SheetTextBox {
        anchor_row: 1,
        x_offset_pt: 120.0,
        y_offset_pt: 0.0,
        width: 40.0,
        height: 20.0,
        paragraphs: vec![Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![Run {
                text: "floating".to_string(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        }],
        fill: None,
        border: None,
        vertical_center: false,
    });
    let source = generate_typst(&make_doc(vec![Page::Sheet(sheet)]))
        .unwrap()
        .source;

    // The grid takes the inset from a `#pad`, but the drawing layer floats in
    // the page foreground, outside that flow, so it has to carry the same
    // inset in its own offsets (issue #1168).
    let inset_pt: f64 = sheet_centering_inset_pt(&source)
        .unwrap_or_else(|| panic!("a centred sheet must inset its grid: {source}"));
    let dx: String = format!("dx: {}pt", 50.0 + inset_pt + 120.0);
    assert!(
        source.contains(&dx),
        "the drawing must move with the centred grid ({dx}): {source}"
    );
}

/// A worksheet line shape floats in the drawing layer like a picture does,
/// and paints as a stroked line at its anchor's page offset: the budget
/// workbook's separators are `a:ln w="12700"` connectors traced by Excel as
/// 0.78pt #D9D9D9 lines on the fitted page (issue #1566).
#[test]
fn test_sheet_line_shape_paints_as_a_stroked_line_in_the_drawing_layer() {
    use crate::ir::{ArrowHead, BorderLineStyle, BorderSide, Color, LineJoin, Shape, ShapeKind};

    let Page::Sheet(mut sheet) = centered_sheet_page(false, vec![60.0; 5]) else {
        unreachable!("centered_sheet_page builds a sheet page")
    };
    sheet.shapes.push(crate::ir::SheetShape {
        anchor_row: 4,
        x_offset_pt: 120.0,
        y_offset_pt: 30.5,
        width: 0.0,
        height: 154.05,
        shape: Shape {
            kind: ShapeKind::Line {
                x1: 0.0,
                y1: 0.0,
                x2: 0.0,
                y2: 154.05,
                head_end: ArrowHead::None,
                tail_end: ArrowHead::None,
            },
            fill: None,
            gradient_fill: None,
            pattern_fill: None,
            stroke: Some(BorderSide {
                width: 0.78,
                color: Color::new(217, 217, 217),
                style: BorderLineStyle::Solid,
                join: LineJoin::Round,
            }),
            rotation_deg: None,
            opacity: None,
            shadow: None,
            top_bevel: None,
        },
    });
    let source = generate_typst(&make_doc(vec![Page::Sheet(sheet)]))
        .unwrap()
        .source;

    assert!(
        source.contains(", foreground: "),
        "a sheet whose only drawing is a line still floats a drawing layer: {source}"
    );
    // Offsets are page-relative: the 50pt left and 54pt top margins fold in.
    assert!(
        source.contains("#place(top + left, dy: 84.5pt)[#place(top + left, dx: 170pt)["),
        "the line is placed at its anchor's page offset: {source}"
    );
    assert!(
        source.contains(
            "#line(start: (0pt, 0pt), end: (0pt, 154.05pt), stroke: (paint: rgb(217, 217, 217), thickness: 0.78pt, join: \"round\"))"
        ),
        "the line is stroked with its declared width and colour: {source}"
    );
}

/// A sheet whose grid is one filled panel, with a picture anchored inside it.
///
/// Modelled on the reported workbook's `Gift budget and tracker` sheet, whose
/// photo sits inside a pale `#F8F0F1` panel (issue #1168).
#[cfg(not(target_arch = "wasm32"))]
fn sheet_with_a_picture_over_a_filled_panel() -> Page {
    use crate::ir::{Color, ImageData, ImageFormat, SheetImage};

    const PANEL: Color = Color {
        r: 0xF8,
        g: 0xF0,
        b: 0xF1,
    };
    /// A 1x1 opaque PNG, enough for the layout engine to place a picture.
    const PIXEL_PNG: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08, 0xD7, 0x63, 0xF8,
        0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xDD, 0x8D, 0xB0, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    let panel_row = |height_pt: f64| TableRow {
        minimum_height: None,
        height: Some(height_pt),
        cells: vec![TableCell {
            content: Vec::new(),
            background: Some(PANEL),
            ..TableCell::default()
        }],
    };

    Page::Sheet(SheetPage {
        name: "Gift budget and tracker".to_string(),
        size: PageSize::default(),
        margins: Margins::default(),
        table: Table {
            rows: vec![panel_row(200.0), panel_row(40.0), panel_row(40.0)],
            column_widths: vec![200.0],
            ..Table::default()
        },
        header: None,
        footer: None,
        charts: Vec::new(),
        images: vec![SheetImage {
            anchor_row: 1,
            x_offset_pt: 20.0,
            y_offset_pt: 10.0,
            clip_width_pt: None,
            image: ImageData {
                data: PIXEL_PNG.to_vec(),
                format: ImageFormat::Png,
                width: Some(120.0),
                height: Some(90.0),
                rotation_deg: None,
                flip_h: false,
                flip_v: false,
                crop: None,
                stroke: None,
                alignment: None,
                clip_shape: None,
                shadow: None,
                paragraph_spacing: None,
            },
        }],
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    })
}

/// Excel floats a drawing above the cells, so a picture anchored inside a
/// filled panel stays visible. Painting the drawing before the grid put every
/// cell fill on top of it and the picture disappeared entirely (issue #1168).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_anchored_sheet_picture_paints_above_the_cell_fills() {
    use crate::render::pdf::{PaintedKind, compiled_paint_sequence};

    let doc = make_doc(vec![sheet_with_a_picture_over_a_filled_panel()]);
    let output = generate_typst(&doc).unwrap();
    let painted =
        compiled_paint_sequence(&output.source, &output.images, 0).expect("the sheet compiles");

    let picture_at: usize = painted
        .iter()
        .position(|item| item.kind == PaintedKind::Image)
        .unwrap_or_else(|| panic!("the sheet's picture is painted: {painted:?}"));
    let picture = painted[picture_at].clone();
    let covering: Vec<usize> = painted
        .iter()
        .enumerate()
        .filter(|(index, item)| {
            *index != picture_at && item.kind == PaintedKind::Shape && item.covers(&picture)
        })
        .map(|(index, _)| index)
        .collect();

    // Without a fill that reaches over the picture the ordering is untestable,
    // so the panel's coverage is asserted before the order it is painted in.
    assert!(
        !covering.is_empty(),
        "the panel fills must cover the picture's box for this to test anything: {painted:?}"
    );
    assert!(
        covering.iter().all(|index| *index < picture_at),
        "cell fills at {covering:?} paint over the picture at {picture_at}: {painted:?}"
    );
}

/// The same sheet, with `row_count` panel rows of `row_height_pt` each, so a
/// tall one breaks across printed pages the way Typst paginates any grid.
#[cfg(not(target_arch = "wasm32"))]
fn sheet_with_a_picture_over_rows(row_count: usize, row_height_pt: f64) -> Page {
    let Page::Sheet(mut sheet) = sheet_with_a_picture_over_a_filled_panel() else {
        unreachable!("sheet_with_a_picture_over_a_filled_panel builds a sheet page")
    };
    let row = sheet.table.rows[1].clone();
    sheet.table.rows = std::iter::repeat_n(row, row_count)
        .map(|mut row| {
            row.height = Some(row_height_pt);
            row
        })
        .collect();
    Page::Sheet(sheet)
}

/// Excel prints a drawing on the page its anchor sits on. A sheet taller than
/// one page breaks across regions, and a `#place` following the grid resolves
/// against the last of them — which is why the drawings float in the page
/// foreground rather than simply being emitted after the grid (issue #1168).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_a_sheet_taller_than_one_page_keeps_its_drawing_on_the_first() {
    use crate::render::pdf::{PaintedKind, compiled_paint_sequence};

    let doc = make_doc(vec![sheet_with_a_picture_over_rows(20, 100.0)]);
    let output = generate_typst(&doc).unwrap();
    let pictures_on = |page_index: usize| -> usize {
        compiled_paint_sequence(&output.source, &output.images, page_index)
            .unwrap_or_else(|error| panic!("page {page_index} compiles: {error}"))
            .iter()
            .filter(|item| item.kind == PaintedKind::Image)
            .count()
    };

    assert_eq!(
        pictures_on(0),
        1,
        "the anchored picture prints on the sheet's first page"
    );
    assert_eq!(
        pictures_on(1),
        0,
        "the picture must not repeat on, or move to, a continuation page"
    );
}

/// A `#set page` rule carries forward, so a sheet that declares no drawings
/// still inherits the previous sheet's foreground. It must draw nothing there:
/// the layer recognises its own sheet by the marker in its content (#1168).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn test_a_later_sheet_does_not_repeat_the_previous_sheets_drawings() {
    use crate::render::pdf::{PaintedKind, compiled_paint_sequence};

    let Page::Sheet(mut plain) = sheet_with_a_picture_over_a_filled_panel() else {
        unreachable!("sheet_with_a_picture_over_a_filled_panel builds a sheet page")
    };
    plain.name = "Data".to_string();
    plain.images.clear();

    let doc = make_doc(vec![
        sheet_with_a_picture_over_a_filled_panel(),
        Page::Sheet(plain),
    ]);
    let output = generate_typst(&doc).unwrap();
    let pictures_on = |page_index: usize| -> usize {
        compiled_paint_sequence(&output.source, &output.images, page_index)
            .unwrap_or_else(|error| panic!("page {page_index} compiles: {error}"))
            .iter()
            .filter(|item| item.kind == PaintedKind::Image)
            .count()
    };

    assert_eq!(pictures_on(0), 1, "the first sheet keeps its picture");
    assert_eq!(
        pictures_on(1),
        0,
        "a sheet with no drawings of its own prints none"
    );
}

/// Restate the drawing geometry of the public #982 workbook without making
/// the regression depend on an installed Segoe UI face. The original fits at
/// 82%; its 100% control splits at a 1,070pt column group. Both native exports
/// keep the same declared chart and picture dimensions before scaling.
#[cfg(not(target_arch = "wasm32"))]
fn gift_drawing_origin_probe(scale: f64) -> Vec<crate::render::pdf::PaintedPrimitive> {
    gift_drawing_origin_probe_with_grid_scale(scale, true)
}

#[cfg(not(target_arch = "wasm32"))]
fn gift_drawing_origin_probe_with_grid_scale(
    scale: f64,
    fitted_grid: bool,
) -> Vec<crate::render::pdf::PaintedPrimitive> {
    gift_drawing_origin_probe_with_chart(scale, fitted_grid, |_| {})
}

#[cfg(not(target_arch = "wasm32"))]
fn gift_drawing_origin_probe_with_chart(
    scale: f64,
    fitted_grid: bool,
    configure: impl FnOnce(&mut Chart),
) -> Vec<crate::render::pdf::PaintedPrimitive> {
    let mut sheet = sheet_page_with_chart_print_scale(scale);
    sheet.size = PageSize {
        width: 1_190.55,
        height: 841.89,
    };
    sheet.margins = Margins {
        top: 54.0,
        bottom: 54.0,
        left: 50.0,
        right: 50.0,
    };
    sheet.table = Table {
        rows: vec![TableRow {
            cells: vec![TableCell::default()],
            height: Some(20.0),
            minimum_height: None,
        }],
        column_widths: vec![if scale == 1.0 { 1_070.0 } else { 1_078.3 }],
        centers_between_print_margins: true,
        print_scale: (fitted_grid && scale < 1.0).then_some(scale),
        ..Table::default()
    };
    let chart = &mut sheet.charts[0];
    chart.placement = Some(crate::ir::SheetChartPlacement {
        x_offset_pt: 285.9874 * scale,
        y_offset_pt: 78.0126 * scale,
        width: 1_015.978_4,
        height: 307.9732,
        print_scale: scale,
        clip_window: None,
    });
    chart.chart.host = crate::ir::ChartHost::Spreadsheet;
    chart.chart.chart_type = ChartType::Column;
    chart.chart.text_font_family = Some("Calibri".to_string());
    chart.chart.text_style.size_pt = Some(9.0);
    chart.chart.categories = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ]
    .into_iter()
    .map(str::to_string)
    .collect();
    chart.chart.series[0].values = vec![
        30.0, 0.0, 0.0, 20.0, 0.0, 180.0, 70.0, 0.0, 0.0, 0.0, 0.0, 0.0,
    ];
    chart.chart.title = None;
    chart.chart.has_legend = true;
    chart.chart.legend_position = LegendPosition::Bottom;
    chart.chart.auto_title_deleted = true;
    chart.chart.chart_area_fill = crate::ir::ChartAreaFill::Solid(Color {
        r: 255,
        g: 255,
        b: 255,
    });
    chart.chart.chart_area_outline = ChartAreaOutline::Suppressed;
    configure(&mut chart.chart);

    let Page::Sheet(picture_sheet) = sheet_with_a_picture_over_a_filled_panel() else {
        unreachable!("the picture helper builds a sheet")
    };
    sheet.images = picture_sheet.images;
    let picture = &mut sheet.images[0];
    picture.x_offset_pt = 10.935437 * scale;
    picture.y_offset_pt = 467.40316 * scale;
    picture.image.width = Some(234.94803 * scale);
    picture.image.height = Some(171.0504 * scale);

    let output = generate_typst(&make_doc(vec![Page::Sheet(sheet)])).unwrap();
    crate::render::pdf::compiled_paint_sequence(&output.source, &output.images, 0)
        .expect("the drawing-origin probe compiles")
}

#[cfg(not(target_arch = "wasm32"))]
fn assert_gift_drawing_bounds(
    scale: f64,
    expected_chart: (f64, f64, f64, f64),
    expected_image: (f64, f64, f64, f64),
) {
    use crate::render::pdf::PaintedKind;

    let painted = gift_drawing_origin_probe(scale);
    let chart = painted
        .iter()
        .find(|item| {
            item.kind == PaintedKind::Shape
                && ((item.bounds.2 - item.bounds.0) - 1_015.978_4 * scale).abs() < 0.01
                && ((item.bounds.3 - item.bounds.1) - 307.9732 * scale).abs() < 0.01
        })
        .expect("the chart paints its full declared frame");
    let image = painted
        .iter()
        .find(|item| item.kind == PaintedKind::Image)
        .expect("the anchored image is present");
    for (name, actual, expected) in [
        ("chart", chart.bounds, expected_chart),
        ("image", image.bounds, expected_image),
    ] {
        for (edge, actual, expected) in [
            ("left", actual.0, expected.0),
            ("top", actual.1, expected.1),
            ("right", actual.2, expected.2),
            ("bottom", actual.3, expected.3),
        ] {
            assert!(
                (actual - expected).abs() < 0.01,
                "{name} {edge} at {scale}: got {actual}pt, expected {expected}pt"
            );
        }
    }
}

/// Fresh native Excel page-2 frame/image bounds for issue #1542. This checks
/// compiled page coordinates, so moving a grid-only marker cannot satisfy it.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn fitted_drawing_frames_use_the_native_sheet_origin() {
    assert_gift_drawing_bounds(
        0.82,
        (289.4497, 117.2703, 1_122.552_0, 369.8083),
        (63.90705, 436.57063, 256.56445, 576.83193),
    );
}

/// Freeze the existing 100% control. Its 0.275pt horizontal native difference
/// is below the visual gate and is independent of the fitted-origin defect.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn unscaled_drawing_frames_keep_their_existing_origin() {
    assert_gift_drawing_bounds(
        1.0,
        (345.2624, 132.0126, 1_361.240_8, 439.9858),
        (70.210437, 521.40316, 305.158467, 692.45356),
    );
}

/// Native Excel for Mac 16.112 exports of the gift workbook put the column
/// plot's top gridline 11 sheet points below the chart frame and its bottom
/// axis rule 48.793 sheet points (11 + the 9pt category and legend bands)
/// above the frame's bottom edge, at the fitted 0.82 print scale (frame top
/// 117.2703pt, rules 126.2903..329.7977pt) and at the unscaled control (frame
/// top 132.0126pt, rules 143.0126..391.1924pt) alike. The chart content
/// therefore shares the frame's fitted sheet origin rather than keeping the
/// converter's physical one, which sat 0.854 sheet points lower at 0.82 and
/// left the unscaled plot 0.854pt high (issue #1607).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn worksheet_column_plot_edges_follow_native_at_every_print_scale() {
    for (scale, expected_top, expected_bottom) in
        [(1.0, 143.0126, 391.1924), (0.82, 126.2903, 329.7977)]
    {
        let rules: Vec<f64> = gift_drawing_origin_probe(scale)
            .into_iter()
            .filter(|item| {
                item.stroke.is_some()
                    && item.bounds.2 - item.bounds.0 > 700.0 * scale
                    && item.bounds.3 - item.bounds.1 < 1.0
            })
            .map(|item| (item.bounds.1 + item.bounds.3) / 2.0)
            .collect();
        assert!(
            rules.len() >= 11,
            "at {scale}: expected the plot's value rules, got {rules:?}"
        );
        let top: f64 = rules.iter().copied().fold(f64::INFINITY, f64::min);
        let bottom: f64 = rules.iter().copied().fold(f64::NEG_INFINITY, f64::max);
        assert!(
            (top - expected_top).abs() < 0.02,
            "plot top rule at {scale}: got {top}pt, native {expected_top}pt"
        );
        assert!(
            (bottom - expected_bottom).abs() < 0.02,
            "plot bottom rule at {scale}: got {bottom}pt, native {expected_bottom}pt"
        );
    }
}

/// The translation the fitted sheet origin applies to the chart frame at
/// `scale`, read off the compiled frame itself: its bounds with the grid on
/// the physical origin against the same grid on the fitted one.
#[cfg(not(target_arch = "wasm32"))]
fn fitted_frame_shift(scale: f64) -> (f64, f64) {
    use crate::render::pdf::PaintedKind;

    let frame_origin = |fitted_grid: bool| {
        gift_drawing_origin_probe_with_grid_scale(scale, fitted_grid)
            .into_iter()
            .find(|item| {
                item.kind == PaintedKind::Shape
                    && ((item.bounds.2 - item.bounds.0) - 1_015.978_4 * scale).abs() < 0.01
                    && ((item.bounds.3 - item.bounds.1) - 307.9732 * scale).abs() < 0.01
            })
            .map(|item| (item.bounds.0, item.bounds.1))
            .expect("the chart paints its full declared frame")
    };
    let physical = frame_origin(false);
    let fitted = frame_origin(true);
    let shift = (fitted.0 - physical.0, fitted.1 - physical.1);
    assert!(
        shift.0.abs() > 0.05 || shift.1.abs() > 0.05,
        "the probe must move the frame onto a distinct fitted origin: {shift:?}"
    );
    shift
}

/// Check that every text run of a fitted chart followed its frame onto the
/// fitted sheet origin: a run either moves by exactly the frame's shift, or
/// it is value-axis chrome that Excel seats on absolute whole sheet points
/// (#1471) and so moves by a whole number of printed sheet points instead.
/// Returns how many runs moved with the frame exactly.
#[cfg(not(target_arch = "wasm32"))]
fn assert_text_runs_follow_the_frame(
    context: &str,
    scale: f64,
    physical: &[(f64, f64, f64, f64)],
    fitted: &[(f64, f64, f64, f64)],
    (dx, dy): (f64, f64),
) -> usize {
    assert!(
        !physical.is_empty(),
        "{context}: the chart must contain text"
    );
    assert_eq!(
        physical.len(),
        fitted.len(),
        "{context}: the runs must not reflow"
    );
    let mut exact: usize = 0;
    for (index, (before, after)) in physical.iter().zip(fitted).enumerate() {
        let moved_sheet_pt: f64 = (after.1 - before.1) / scale;
        let with_frame: bool = (after.1 - before.1 - dy).abs() < 0.01;
        let on_whole_sheet_points: bool =
            (moved_sheet_pt - moved_sheet_pt.round()).abs() < 0.01 && moved_sheet_pt.abs() <= 2.0;
        assert!(
            (after.0 - before.0 - dx).abs() < 0.01 && (with_frame || on_whole_sheet_points),
            "{context}, run {index}: did not follow the frame's ({dx}, {dy})pt shift: {before:?} -> {after:?}"
        );
        exact += usize::from(with_frame);
    }
    assert!(
        exact > 0,
        "{context}: no run followed the frame exactly, so the shift is not the frame's"
    );
    exact
}

/// Excel lays a chart out inside the frame it paints, so the fitted origin
/// correction of #1542 carries the chart's text with the frame: every run
/// moves by the frame's own shift, none stays on the physical origin
/// (issue #1607).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn fitted_drawing_frames_carry_chart_text_with_the_frame() {
    use crate::render::pdf::PaintedKind;

    let text_bounds = |fitted_grid| {
        gift_drawing_origin_probe_with_grid_scale(0.82, fitted_grid)
            .into_iter()
            .filter(|item| item.kind == PaintedKind::Text)
            .map(|item| item.bounds)
            .collect::<Vec<_>>()
    };
    let shift = fitted_frame_shift(0.82);
    let physical = text_bounds(false);
    let fitted = text_bounds(true);
    assert_text_runs_follow_the_frame("gift chart at 0.82", 0.82, &physical, &fitted, shift);
    // The category labels and the legend sit under the plot's bottom rule
    // (329.80pt on the page) and are not snapped, so they carry the frame's
    // exact shift; only the value labels above it may re-round.
    for (index, (before, after)) in physical.iter().zip(&fitted).enumerate() {
        if before.1 < 330.0 {
            continue;
        }
        assert!(
            (after.1 - before.1 - shift.1).abs() < 0.01,
            "run {index} below the plot did not follow the frame's {}pt shift: {before:?} -> {after:?}",
            shift.1
        );
    }
}

/// The plot's top and bottom rules are laid out from the frame's edges, so
/// they follow the frame onto the fitted origin by exactly its shift. The
/// interior gridlines snap to whole sheet points on that origin and are
/// pinned against native by `worksheet_column_plot_edges_follow_native_at_every_print_scale`
/// and the sheet-space snapping tests (issue #1607).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn fitted_drawing_frames_carry_the_plot_edges_with_the_frame() {
    let plot_edges = |fitted_grid| {
        let rules: Vec<(f64, f64)> = gift_drawing_origin_probe_with_grid_scale(0.82, fitted_grid)
            .into_iter()
            .filter(|item| {
                item.stroke.is_some()
                    && item.bounds.2 - item.bounds.0 > 700.0
                    && item.bounds.3 - item.bounds.1 < 1.0
            })
            .map(|item| (item.bounds.0, (item.bounds.1 + item.bounds.3) / 2.0))
            .collect();
        assert!(rules.len() >= 5, "the probe must contain plot gridlines");
        let top = rules
            .iter()
            .copied()
            .min_by(|a, b| a.1.total_cmp(&b.1))
            .expect("a top rule");
        let bottom = rules
            .iter()
            .copied()
            .max_by(|a, b| a.1.total_cmp(&b.1))
            .expect("a bottom rule");
        (top, bottom)
    };
    let (dx, dy) = fitted_frame_shift(0.82);
    let (before_top, before_bottom) = plot_edges(false);
    let (after_top, after_bottom) = plot_edges(true);
    for (name, before, after) in [
        ("top", before_top, after_top),
        ("bottom", before_bottom, after_bottom),
    ] {
        assert!(
            (after.0 - before.0 - dx).abs() < 0.01 && (after.1 - before.1 - dy).abs() < 0.01,
            "plot {name} rule did not follow the frame's ({dx}, {dy})pt shift: {before:?} -> {after:?}"
        );
    }
}

/// The bottom legend's filled key is seated from the frame's bottom edge and
/// follows it onto the fitted origin like the rest of the content.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn fitted_drawing_frames_carry_the_legend_key_with_the_frame() {
    use crate::render::pdf::PaintedKind;

    let key = |fitted_grid| {
        gift_drawing_origin_probe_with_grid_scale(0.82, fitted_grid)
            .into_iter()
            .find(|item| {
                item.kind == PaintedKind::Shape
                    && item.stroke.is_none()
                    && item.bounds.1 > 340.0
                    && item.bounds.2 - item.bounds.0 < 25.0
                    && item.bounds.3 - item.bounds.1 < 10.0
            })
            .expect("the bottom legend contains a filled key")
            .bounds
    };
    let (dx, dy) = fitted_frame_shift(0.82);
    let before = key(false);
    let after = key(true);
    for (before, after, offset) in [
        (before.0, after.0, dx),
        (before.1, after.1, dy),
        (before.2, after.2, dx),
        (before.3, after.3, dy),
    ] {
        assert!(
            (after - before - offset).abs() < 0.01,
            "legend key edge did not follow the frame's shift {offset}pt: {before} -> {after}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn fitted_drawing_frames_preserve_chart_text_flow_at_other_scales() {
    use crate::render::pdf::PaintedKind;

    for scale in [0.64, 0.78] {
        for chart_type in [ChartType::Column, ChartType::Bar, ChartType::Line] {
            let bounds = |fitted_grid| {
                gift_drawing_origin_probe_with_chart(scale, fitted_grid, |chart| {
                    chart.chart_type = chart_type.clone();
                    chart.title = Some("Budget by month\nPlanned and actual spending".to_string());
                    chart.auto_title_deleted = false;
                    chart.series[0].name = Some("Planned spending\nCurrent year".to_string());
                    chart.categories[0] = "Beginning of January".to_string();
                })
                .into_iter()
                .filter(|item| item.kind == PaintedKind::Text)
                .map(|item| item.bounds)
                .collect::<Vec<_>>()
            };
            let shift = fitted_frame_shift(scale);
            let _ = assert_text_runs_follow_the_frame(
                &format!("{chart_type:?} at {scale}"),
                scale,
                &bounds(false),
                &bounds(true),
                shift,
            );
        }
    }
}
