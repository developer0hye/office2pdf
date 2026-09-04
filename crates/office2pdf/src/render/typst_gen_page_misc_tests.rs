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
                frame: None,
            }],
        }),
        footer: None,
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
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
                frame: None,
            }],
        }),
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
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
#[cfg(not(target_arch = "wasm32"))]
fn sheet_page_with_seated_footer(distance_from_edge: Option<f64>, family: Option<&str>) -> Page {
    Page::Sheet(SheetPage {
        name: "Sheet1".to_string(),
        size: PageSize::default(),
        margins: Margins {
            top: 54.0,
            bottom: 54.0,
            left: 50.0,
            right: 50.0,
        },
        table: make_simple_table(vec![vec!["A"]]),
        header: None,
        footer: Some(HeaderFooter {
            shapes: Vec::new(),
            distance_from_edge,
            sheet_print_scale: None,
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle {
                    alignment: Some(Alignment::Left),
                    ..ParagraphStyle::default()
                },
                elements: vec![HFInline::Run(Run {
                    text: "Sensitivity: Internal".to_string(),
                    style: TextStyle {
                        font_family: family.map(str::to_string),
                        font_size: Some(8.0),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                })],
                border: None,
                border_space: None,
                frame: None,
            }],
        }),
        charts: vec![],
        images: Vec::new(),
        text_boxes: Vec::new(),
    })
}

/// Excel lays a fitted sheet's header/footer out in sheet coordinates and
/// scales that box onto the paper. At 0.82, the A3 probe's 50pt page margins
/// therefore become the outward-rounded box 49.2..1141.44pt (#1510).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_fitted_sheet_footer_uses_its_scaled_horizontal_coordinate_box() {
    let mut page = sheet_page_with_seated_footer(Some(23.0), Some("Arial"));
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
        Some(23.0),
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
        "the band must span the 54pt bottom margin down to the 23pt seat: {source}"
    );
    assert!(
        source.contains("#place(bottom"),
        "the content must rest on the band's bottom: {source}"
    );
}

/// Triangulation: the band is measured, not a constant (issue #1142).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_seated_sheet_footer_band_tracks_the_stated_seat() {
    let source = generate_typst(&make_doc(vec![sheet_page_with_seated_footer(
        Some(38.0),
        Some("Arial"),
    )]))
    .expect("document should generate")
    .source;

    assert!(
        source.contains("block(width: 100%, height: 16pt)"),
        "a 38pt seat under a 54pt bottom margin leaves a 16pt band: {source}"
    );
}

/// The band states the footer face's own `hhea` descent, so the last baseline
/// lands the descent above the seat (issue #1142).
///
/// Arial, because every runner resolves it — through Liberation Sans where the
/// face itself is absent.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_seated_sheet_footer_states_its_own_bottom_edge() {
    let (ascender_em, _, pitch_em) =
        crate::render::pdf::font_line_metrics_em("Arial").expect("Arial metrics should resolve");
    let expected_em: f64 = pitch_em - ascender_em;

    let source = generate_typst(&make_doc(vec![sheet_page_with_seated_footer(
        Some(23.0),
        Some("Arial"),
    )]))
    .expect("document should generate")
    .source;

    assert!(
        source.contains(&format!("bottom-edge: -{}em", format_f64(expected_em))),
        "the band must state Arial's own {expected_em}em descent: {source}"
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
    });
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    assert!(!output.source.contains("header:"));
    assert!(!output.source.contains("footer:"));
}

#[test]
fn test_table_page_with_anchored_chart_overlays_the_grid() {
    use crate::ir::{Chart, ChartGrouping, ChartSeries, ChartType, DataLabels, LegendPosition};

    let chart = Chart {
        chart_type: ChartType::Bar,
        hole_size_percent: None,
        title: Some("Sales".to_string()),
        categories: vec!["Q1".to_string(), "Q2".to_string()],
        series: vec![ChartSeries {
            name: Some("Revenue".to_string()),
            values: vec![100.0, 200.0],
            fill: None,
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            marker_symbol: None,
            line_width_pt: None,
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
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_number_format: None,
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
            }),
            chart,
        }],
        images: Vec::new(),
        text_boxes: Vec::new(),
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
    use crate::ir::{Chart, ChartGrouping, ChartSeries, ChartType, DataLabels, LegendPosition};

    let chart = Chart {
        chart_type: ChartType::Bar,
        hole_size_percent: None,
        title: Some("Sales".to_string()),
        categories: vec!["Q1".to_string(), "Q2".to_string()],
        series: vec![ChartSeries {
            name: Some("Revenue".to_string()),
            values: vec![100.0, 200.0],
            fill: None,
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            marker_symbol: None,
            line_width_pt: None,
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
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_number_format: None,
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
            }),
            chart,
        }],
        images: Vec::new(),
        text_boxes: Vec::new(),
    });

    generate_typst(&make_doc(vec![page])).unwrap().source
}

#[test]
fn test_table_page_with_chart_at_end() {
    use crate::ir::{Chart, ChartGrouping, ChartSeries, ChartType, DataLabels, LegendPosition};

    let chart = Chart {
        chart_type: ChartType::Pie,
        hole_size_percent: None,
        title: Some("Pie".to_string()),
        categories: vec!["A".to_string()],
        series: vec![ChartSeries {
            name: None,
            values: vec![100.0],
            fill: None,
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            marker_symbol: None,
            line_width_pt: None,
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
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_number_format: None,
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
            "gradient.linear((rgb(159, 223, 191), 0%), (rgb(78, 179, 207), 100%), angle: 32deg)"
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
    let picture = painted[picture_at];
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
