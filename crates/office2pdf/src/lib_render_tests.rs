use super::test_support::{make_simple_document, make_test_png};
use super::*;
use crate::ir::*;

#[test]
fn test_render_document_empty_document() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![],
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(!pdf.is_empty(), "PDF bytes should not be empty");
    assert!(pdf.starts_with(b"%PDF"), "Should be valid PDF");
}

#[test]
fn test_render_document_single_paragraph() {
    let doc = make_simple_document("Hello, World!");
    let pdf = render_document(&doc).unwrap();
    assert!(!pdf.is_empty());
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_render_document_with_tab_leader() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
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
            })],
            header: None,
            footer: None,
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };

    let pdf = render_document(&doc).unwrap();
    assert!(!pdf.is_empty());
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_render_document_styled_text() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle {
                    alignment: Some(Alignment::Center),
                    ..ParagraphStyle::default()
                },
                runs: vec![
                    Run {
                        text: "Bold text ".to_string(),
                        style: TextStyle {
                            bold: Some(true),
                            font_size: Some(16.0),
                            ..TextStyle::default()
                        },
                        href: None,
                        footnote: None,
                    },
                    Run {
                        text: "and italic".to_string(),
                        style: TextStyle {
                            italic: Some(true),
                            color: Some(Color::new(255, 0, 0)),
                            ..TextStyle::default()
                        },
                        href: None,
                        footnote: None,
                    },
                ],
            })],
            header: None,
            footer: None,
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(!pdf.is_empty());
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_render_document_multiple_flow_pages() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![
            Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Page 1".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                })],
                header: None,
                footer: None,
                page_number_start: None,
                columns: None,
            }),
            Page::Flow(FlowPage {
                size: PageSize::default(),
                margins: Margins::default(),
                content: vec![Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Page 2".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                })],
                header: None,
                footer: None,
                page_number_start: None,
                columns: None,
            }),
        ],
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(!pdf.is_empty());
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_render_document_page_break() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![
                Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Before break".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }),
                Block::PageBreak,
                Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "After break".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }),
            ],
            header: None,
            footer: None,
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(!pdf.is_empty());
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_render_document_with_image() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Image(ImageData {
                data: make_test_png(),
                format: ImageFormat::Png,
                width: Some(100.0),
                height: Some(80.0),
                crop: None,
            })],
            header: None,
            footer: None,
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(!pdf.is_empty(), "PDF should not be empty");
    assert!(pdf.starts_with(b"%PDF"), "Should be valid PDF");
}

#[test]
fn test_render_document_image_mixed_with_text() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![
                Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Image below:".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }),
                Block::Image(ImageData {
                    data: make_test_png(),
                    format: ImageFormat::Png,
                    width: Some(200.0),
                    height: None,
                    crop: None,
                }),
                Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Image above.".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }),
            ],
            header: None,
            footer: None,
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(!pdf.is_empty());
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_render_document_with_system_font_in_ir() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Hello with system font".to_string(),
                    style: TextStyle {
                        font_family: Some("Arial".to_string()),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                }],
            })],
            header: None,
            footer: None,
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(!pdf.is_empty());
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_render_document_with_multiple_font_families() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![
                    Run {
                        text: "Calibri text ".to_string(),
                        style: TextStyle {
                            font_family: Some("Calibri".to_string()),
                            ..TextStyle::default()
                        },
                        href: None,
                        footnote: None,
                    },
                    Run {
                        text: "and Times New Roman text".to_string(),
                        style: TextStyle {
                            font_family: Some("Times New Roman".to_string()),
                            ..TextStyle::default()
                        },
                        href: None,
                        footnote: None,
                    },
                ],
            })],
            header: None,
            footer: None,
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(!pdf.is_empty());
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_render_document_with_list() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::List(List {
                kind: ListKind::Unordered,
                items: vec![
                    ListItem {
                        content: vec![Paragraph {
                            style: ParagraphStyle::default(),
                            runs: vec![Run {
                                text: "Hello".to_string(),
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
                                text: "World".to_string(),
                                style: TextStyle::default(),
                                href: None,
                                footnote: None,
                            }],
                        }],
                        level: 0,
                        start_at: None,
                    },
                ],
                level_styles: std::collections::BTreeMap::new(),
            })],
            header: None,
            footer: None,
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(
        pdf.starts_with(b"%PDF"),
        "Should produce valid PDF with list"
    );
}

#[test]
fn test_render_document_with_header() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Body content".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
            header: Some(HeaderFooter {
                paragraphs: vec![HeaderFooterParagraph {
                    style: ParagraphStyle::default(),
                    elements: vec![HFInline::Run(Run {
                        text: "My Header".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    })],
                }],
            }),
            footer: None,
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(!pdf.is_empty());
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_render_document_with_page_number_footer() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Body content".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            })],
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
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(!pdf.is_empty());
    assert!(pdf.starts_with(b"%PDF"));
}

#[test]
fn test_render_document_with_landscape_page() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Flow(FlowPage {
            size: PageSize {
                width: 841.9,
                height: 595.3,
            },
            margins: Margins::default(),
            content: vec![Block::Paragraph(Paragraph {
                runs: vec![Run {
                    text: "Landscape page".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
                style: ParagraphStyle::default(),
            })],
            header: None,
            footer: None,
            page_number_start: None,
            columns: None,
        })],
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(
        !pdf.is_empty(),
        "Landscape FlowPage should produce non-empty PDF"
    );
    assert!(pdf.starts_with(b"%PDF"), "Should produce valid PDF");
}

#[test]
fn test_render_multipage_document_size() {
    let mut pages = Vec::new();
    for i in 1..=10 {
        pages.push(Page::Flow(FlowPage {
            size: PageSize::default(),
            margins: Margins::default(),
            content: vec![
                Block::Paragraph(Paragraph {
                    style: ParagraphStyle {
                        alignment: Some(Alignment::Center),
                        ..ParagraphStyle::default()
                    },
                    runs: vec![Run {
                        text: format!("Page {i} Heading"),
                        style: TextStyle {
                            bold: Some(true),
                            font_size: Some(24.0),
                            ..TextStyle::default()
                        },
                        href: None,
                        footnote: None,
                    }],
                }),
                Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: format!(
                            "This is page {i}. Lorem ipsum dolor sit amet, \
                             consectetur adipiscing elit. Sed do eiusmod tempor \
                             incididunt ut labore et dolore magna aliqua."
                        ),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }),
            ],
            header: None,
            footer: None,
            page_number_start: None,
            columns: None,
        }));
    }
    let doc = Document {
        metadata: Metadata::default(),
        pages,
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(
        pdf.len() < 512_000,
        "10-page IR document PDF should be under 500KB, actual: {} bytes ({:.1} KB)",
        pdf.len(),
        pdf.len() as f64 / 1024.0
    );
}

#[test]
fn test_render_pptx_style_document_size() {
    let mut pages = Vec::new();
    for i in 1..=5 {
        pages.push(Page::Fixed(FixedPage {
            size: PageSize {
                width: 720.0,
                height: 540.0,
            },
            background_color: None,
            background_gradient: None,
            elements: vec![FixedElement {
                x: 50.0,
                y: 50.0,
                width: 620.0,
                height: 80.0,
                kind: FixedElementKind::TextBox(TextBoxData {
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: format!("Slide {i} content"),
                            style: TextStyle {
                                font_size: Some(32.0),
                                ..TextStyle::default()
                            },
                            href: None,
                            footnote: None,
                        }],
                    })],
                    padding: Insets::default(),
                    vertical_align: TextBoxVerticalAlign::Top,
                }),
            }],
        }));
    }
    let doc = Document {
        metadata: Metadata::default(),
        pages,
        styles: StyleSheet::default(),
    };
    let pdf = render_document(&doc).unwrap();
    assert!(
        pdf.len() < 512_000,
        "5-slide FixedPage PDF should be under 500KB, actual: {} bytes ({:.1} KB)",
        pdf.len(),
        pdf.len() as f64 / 1024.0
    );
}

#[test]
fn test_render_document_with_centered_fixed_text_box() {
    let doc = Document {
        metadata: Metadata::default(),
        pages: vec![Page::Fixed(FixedPage {
            size: PageSize {
                width: 300.0,
                height: 200.0,
            },
            background_color: None,
            background_gradient: None,
            elements: vec![FixedElement {
                x: 20.0,
                y: 20.0,
                width: 200.0,
                height: 60.0,
                kind: FixedElementKind::TextBox(TextBoxData {
                    content: vec![Block::Paragraph(Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Centered badge".to_string(),
                            style: TextStyle {
                                font_size: Some(18.0),
                                ..TextStyle::default()
                            },
                            href: None,
                            footnote: None,
                        }],
                    })],
                    padding: Insets {
                        top: 3.6,
                        right: 7.2,
                        bottom: 3.6,
                        left: 7.2,
                    },
                    vertical_align: TextBoxVerticalAlign::Center,
                }),
            }],
        })],
        styles: StyleSheet::default(),
    };

    let pdf = render_document(&doc).unwrap();
    assert!(
        pdf.starts_with(b"%PDF"),
        "Centered fixed text box should compile to a valid PDF"
    );
}
