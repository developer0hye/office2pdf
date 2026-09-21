use super::*;

// ----- US-020: Header/footer parsing tests -----

fn build_docx_with_header(header_text: &str) -> Vec<u8> {
    let header = docx_rs::Header::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(header_text)),
    );
    let docx = docx_rs::Docx::new().header(header).add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

fn build_docx_with_footer(footer_text: &str) -> Vec<u8> {
    let footer = docx_rs::Footer::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(footer_text)),
    );
    let docx = docx_rs::Docx::new().footer(footer).add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

fn build_docx_with_page_number_footer() -> Vec<u8> {
    let footer = docx_rs::Footer::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(
            docx_rs::Run::new()
                .add_text("Page ")
                .add_field_char(docx_rs::FieldCharType::Begin, false)
                .add_instr_text(docx_rs::InstrText::PAGE(docx_rs::InstrPAGE::new()))
                .add_field_char(docx_rs::FieldCharType::Separate, false)
                .add_text("1")
                .add_field_char(docx_rs::FieldCharType::End, false),
        ),
    );
    let docx = docx_rs::Docx::new().footer(footer).add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// Word writes a `w:fldChar` field across several runs — one for each of
/// begin, the instruction, separate, the cached result and end. The field's
/// state has to span those runs, or the cached result leaks out as static text
/// and every page shows whatever number was saved (issue #738).
#[test]
fn test_page_field_split_across_runs_is_a_page_number() {
    let footer = docx_rs::Footer::new().add_paragraph(
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_field_char(docx_rs::FieldCharType::Begin, false))
            .add_run(
                docx_rs::Run::new()
                    .add_instr_text(docx_rs::InstrText::PAGE(docx_rs::InstrPAGE::new())),
            )
            .add_run(docx_rs::Run::new().add_field_char(docx_rs::FieldCharType::Separate, false))
            // The number cached when the document was last saved.
            .add_run(docx_rs::Run::new().add_text("7"))
            .add_run(docx_rs::Run::new().add_field_char(docx_rs::FieldCharType::End, false)),
    );
    let docx = docx_rs::Docx::new().footer(footer).add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(page) => page,
        other => panic!("Expected FlowPage, got {other:?}"),
    };
    let footer = page.footer.as_ref().expect("Should have footer");
    let elements: Vec<&crate::ir::HFInline> = footer
        .paragraphs
        .iter()
        .flat_map(|paragraph| paragraph.elements.iter())
        .collect();

    assert!(
        elements
            .iter()
            .any(|element| matches!(element, crate::ir::HFInline::PageNumber(_))),
        "the split field must resolve to a page number, got {elements:?}"
    );
    assert!(
        !elements.iter().any(
            |element| matches!(element, crate::ir::HFInline::Run(run) if run.text.contains("7"))
        ),
        "the cached result must not survive as static text, got {elements:?}"
    );
}

/// A field whose instruction is not one we model still has to show the text
/// Word cached for it. Before the field state spanned runs this happened by
/// accident — `in_field` was false again by the time the result's own run was
/// read — so making the state span runs must not lose it.
#[test]
fn test_unmodelled_split_field_keeps_its_cached_text() {
    let footer = docx_rs::Footer::new().add_paragraph(
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_field_char(docx_rs::FieldCharType::Begin, false))
            .add_run(
                docx_rs::Run::new().add_instr_text(docx_rs::InstrText::Unsupported(
                    " STYLEREF 1 \\s ".to_string(),
                )),
            )
            .add_run(docx_rs::Run::new().add_field_char(docx_rs::FieldCharType::Separate, false))
            .add_run(docx_rs::Run::new().add_text("Chapter Title"))
            .add_run(docx_rs::Run::new().add_field_char(docx_rs::FieldCharType::End, false)),
    );
    let docx = docx_rs::Docx::new().footer(footer).add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(page) => page,
        other => panic!("Expected FlowPage, got {other:?}"),
    };
    let footer = page.footer.as_ref().expect("Should have footer");
    let has_cached_text = footer.paragraphs.iter().any(|paragraph| {
        paragraph.elements.iter().any(
            |element| matches!(element, crate::ir::HFInline::Run(run) if run.text.contains("Chapter Title")),
        )
    });
    assert!(
        has_cached_text,
        "an unmodelled field must still render its cached result"
    );
}

/// A field the paragraph never closes must not swallow the text after it.
#[test]
fn test_unterminated_split_field_does_not_swallow_its_text() {
    let footer = docx_rs::Footer::new().add_paragraph(
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_field_char(docx_rs::FieldCharType::Begin, false))
            .add_run(
                docx_rs::Run::new()
                    .add_instr_text(docx_rs::InstrText::Unsupported(" STYLEREF 1 ".to_string())),
            )
            .add_run(docx_rs::Run::new().add_field_char(docx_rs::FieldCharType::Separate, false))
            .add_run(docx_rs::Run::new().add_text("Dangling")),
    );
    let docx = docx_rs::Docx::new().footer(footer).add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let (doc, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[0] {
        Page::Flow(page) => page,
        other => panic!("Expected FlowPage, got {other:?}"),
    };
    let footer = page.footer.as_ref().expect("Should have footer");
    let has_text = footer.paragraphs.iter().any(|paragraph| {
        paragraph.elements.iter().any(
            |element| matches!(element, crate::ir::HFInline::Run(run) if run.text.contains("Dangling")),
        )
    });
    assert!(has_text, "text inside an unclosed field must still render");
}

/// Word applies the containing run's properties to the field result, so the
/// parsed field must carry that run's style.
#[test]
fn test_page_number_field_carries_its_run_style() {
    let footer = docx_rs::Footer::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(
            docx_rs::Run::new()
                .size(16)
                .color("888888")
                .add_text("- ")
                .add_field_char(docx_rs::FieldCharType::Begin, false)
                .add_instr_text(docx_rs::InstrText::PAGE(docx_rs::InstrPAGE::new()))
                .add_field_char(docx_rs::FieldCharType::Separate, false)
                .add_field_char(docx_rs::FieldCharType::End, false),
        ),
    );
    let docx = docx_rs::Docx::new().footer(footer).add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };
    let elements = &flow.footer.as_ref().expect("footer").paragraphs[0].elements;
    let style = elements
        .iter()
        .find_map(|element| match element {
            crate::ir::HFInline::PageNumber(style) => Some(style),
            _ => None,
        })
        .expect("page number field parsed");
    assert_eq!(style.font_size, Some(8.0), "w:sz 16 half-points is 8pt");
    assert_eq!(style.color, Some(Color::new(0x88, 0x88, 0x88)));
}

fn build_docx_with_total_pages_footer() -> Vec<u8> {
    let footer = docx_rs::Footer::new().add_paragraph(
        docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Total "))
            .add_run(
                docx_rs::Run::new()
                    .add_field_char(docx_rs::FieldCharType::Begin, false)
                    .add_instr_text(docx_rs::InstrText::NUMPAGES(docx_rs::InstrNUMPAGES::new()))
                    .add_field_char(docx_rs::FieldCharType::Separate, false)
                    .add_text("1")
                    .add_field_char(docx_rs::FieldCharType::End, false),
            ),
    );
    let docx = docx_rs::Docx::new()
        .footer(footer)
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body")));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn test_parse_docx_with_text_header() {
    let data = build_docx_with_header("My Document Header");
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    assert!(page.header.is_some(), "FlowPage should have a header");
    let header = page.header.as_ref().unwrap();
    assert!(
        !header.paragraphs.is_empty(),
        "Header should have paragraphs"
    );

    let has_text = header.paragraphs.iter().any(|paragraph| {
        paragraph.elements.iter().any(
            |element| matches!(element, crate::ir::HFInline::Run(run) if run.text.contains("My Document Header")),
        )
    });
    assert!(
        has_text,
        "Header should contain the text 'My Document Header'"
    );
}

#[test]
fn test_parse_docx_with_text_footer() {
    let data = build_docx_with_footer("Footer Text");
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    assert!(page.footer.is_some(), "FlowPage should have a footer");
    let footer = page.footer.as_ref().unwrap();

    let has_text = footer.paragraphs.iter().any(|paragraph| {
        paragraph
            .elements
            .iter()
            .any(|element| matches!(element, crate::ir::HFInline::Run(run) if run.text.contains("Footer Text")))
    });
    assert!(has_text, "Footer should contain 'Footer Text'");
}

#[test]
fn test_parse_docx_with_page_number_in_footer() {
    let data = build_docx_with_page_number_footer();
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    assert!(page.footer.is_some(), "Should have footer");
    let footer = page.footer.as_ref().unwrap();

    let has_page_num = footer.paragraphs.iter().any(|paragraph| {
        paragraph
            .elements
            .iter()
            .any(|element| matches!(element, crate::ir::HFInline::PageNumber(_)))
    });
    assert!(has_page_num, "Footer should contain a PageNumber field");

    let has_text = footer.paragraphs.iter().any(|paragraph| {
        paragraph
            .elements
            .iter()
            .any(|element| matches!(element, crate::ir::HFInline::Run(run) if run.text.contains("Page ")))
    });
    assert!(
        has_text,
        "Footer should contain 'Page ' text before page number"
    );
}

#[test]
fn test_parse_docx_with_total_pages_in_footer() {
    let data = build_docx_with_total_pages_footer();
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    let footer = page.footer.as_ref().expect("Should have footer");
    let has_total_pages = footer.paragraphs.iter().any(|paragraph| {
        paragraph
            .elements
            .iter()
            .any(|element| matches!(element, crate::ir::HFInline::TotalPages(_)))
    });
    assert!(has_total_pages, "Footer should contain a TotalPages field");
}

#[test]
fn test_parse_docx_multiple_sections_with_distinct_page_setup_and_headers() {
    let first_header = docx_rs::Header::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Section One Header")),
    );
    let second_header = docx_rs::Header::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Section Two Header")),
    );

    let first_section = docx_rs::Section::new()
        .page_size(docx_rs::PageSize::new().size(12240, 15840))
        .header(first_header)
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Section One")),
        );

    let docx = docx_rs::Docx::new()
        .add_section(first_section)
        .header(second_header)
        .page_size(15840, 12240)
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Section Two")),
        );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 2, "Expected one FlowPage per DOCX section");

    let first_page = match &doc.pages[0] {
        Page::Flow(page) => page,
        _ => panic!("Expected first page to be FlowPage"),
    };
    let second_page = match &doc.pages[1] {
        Page::Flow(page) => page,
        _ => panic!("Expected second page to be FlowPage"),
    };

    assert!(
        (first_page.size.width - 612.0).abs() < 0.1,
        "first page width should come from first section"
    );
    assert!(
        (first_page.size.height - 792.0).abs() < 0.1,
        "first page height should come from first section"
    );
    assert!(
        (second_page.size.width - 792.0).abs() < 0.1,
        "second page width should come from final section"
    );
    assert!(
        (second_page.size.height - 612.0).abs() < 0.1,
        "second page height should come from final section"
    );

    let first_header_text = first_page
        .header
        .as_ref()
        .and_then(|header_footer| {
            header_footer
                .paragraphs
                .iter()
                .flat_map(|paragraph| paragraph.elements.iter())
                .find_map(|element| match element {
                    crate::ir::HFInline::Run(run) => Some(run.text.as_str()),
                    _ => None,
                })
        })
        .unwrap_or("");
    assert_eq!(first_header_text, "Section One Header");

    let second_header_text = second_page
        .header
        .as_ref()
        .and_then(|header_footer| {
            header_footer
                .paragraphs
                .iter()
                .flat_map(|paragraph| paragraph.elements.iter())
                .find_map(|element| match element {
                    crate::ir::HFInline::Run(run) => Some(run.text.as_str()),
                    _ => None,
                })
        })
        .unwrap_or("");
    assert_eq!(second_header_text, "Section Two Header");
}

#[test]
fn test_parse_docx_with_header_and_footer() {
    let header = docx_rs::Header::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Header Text")),
    );
    let footer = docx_rs::Footer::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Footer Text")),
    );
    let docx = docx_rs::Docx::new()
        .header(header)
        .footer(footer)
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body")));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    assert!(page.header.is_some(), "Should have header");
    assert!(page.footer.is_some(), "Should have footer");
}

#[test]
fn test_parse_docx_without_header_footer() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Just text")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };

    assert!(page.header.is_none(), "No header expected");
    assert!(page.footer.is_none(), "No footer expected");
}

// ----- Page orientation tests -----

#[test]
fn test_portrait_document_width_less_than_height() {
    let data = build_docx_bytes_with_page_setup(
        vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Portrait"))],
        11906,
        16838,
        1440,
        1440,
        1440,
        1440,
    );
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    assert!(
        page.size.width < page.size.height,
        "Portrait: width ({}) should be < height ({})",
        page.size.width,
        page.size.height
    );
}

#[test]
fn test_landscape_document_width_greater_than_height() {
    let data = build_docx_bytes_with_page_setup(
        vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Landscape"))],
        16838,
        11906,
        1440,
        1440,
        1440,
        1440,
    );
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    assert!(
        page.size.width > page.size.height,
        "Landscape: width ({}) should be > height ({})",
        page.size.width,
        page.size.height
    );
    assert!(
        (page.size.width - 841.9).abs() < 1.0,
        "Expected width ~841.9, got {}",
        page.size.width
    );
    assert!(
        (page.size.height - 595.3).abs() < 1.0,
        "Expected height ~595.3, got {}",
        page.size.height
    );
}

#[test]
fn test_default_document_is_portrait() {
    let data = build_docx_bytes(vec![
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Default")),
    ]);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    assert!(
        page.size.width < page.size.height,
        "Default should be portrait: width ({}) < height ({})",
        page.size.width,
        page.size.height
    );
}

#[test]
fn test_landscape_with_orient_attribute() {
    let mut docx = docx_rs::Docx::new()
        .page_size(16838, 11906)
        .page_orient(docx_rs::PageOrientationType::Landscape)
        .page_margin(
            docx_rs::PageMargin::new()
                .top(1440)
                .bottom(1440)
                .left(1440)
                .right(1440),
        );
    docx = docx.add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Landscape with orient")),
    );
    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = match &doc.pages[0] {
        Page::Flow(p) => p,
        _ => panic!("Expected FlowPage"),
    };
    assert!(
        page.size.width > page.size.height,
        "Landscape with orient: width ({}) should be > height ({})",
        page.size.width,
        page.size.height
    );
}

#[test]
fn test_extract_page_size_orient_landscape_swaps_dimensions() {
    let page_size = docx_rs::PageSize::new()
        .width(11906)
        .height(16838)
        .orient(docx_rs::PageOrientationType::Landscape);

    let result = extract_page_size(&page_size);
    assert!(
        result.width > result.height,
        "orient=landscape should ensure width ({}) > height ({})",
        result.width,
        result.height
    );
}

#[test]
fn test_extract_page_size_no_orient_keeps_dimensions() {
    let page_size = docx_rs::PageSize::new().width(11906).height(16838);

    let result = extract_page_size(&page_size);
    assert!(
        result.width < result.height,
        "No orient: width ({}) should be < height ({})",
        result.width,
        result.height
    );
}

/// Word letterhead headers declare the gap between text and rule with
/// `w:pBdr/<side>/@w:space`, in points.
#[test]
fn test_parse_docx_header_paragraph_border_space() {
    let mut paragraph =
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Manual v0.6"));
    paragraph.property = paragraph.property.set_borders(
        docx_rs::ParagraphBorders::with_empty().set(
            docx_rs::ParagraphBorder::new(docx_rs::ParagraphBorderPosition::Bottom)
                .val(docx_rs::BorderType::Single)
                .size(4)
                .space(4)
                .color("CCCCCC"),
        ),
    );
    let header = docx_rs::Header::new().add_paragraph(paragraph);
    let docx = docx_rs::Docx::new().header(header).add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
    );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };
    let paragraph = &flow.header.as_ref().expect("header").paragraphs[0];

    let border = paragraph.border.as_ref().expect("bottom rule parsed");
    assert!(border.bottom.is_some());
    let space = paragraph.border_space.expect("w:space parsed");
    assert_eq!(space.bottom, 4.0);
    assert_eq!(space.top, 0.0);
}

// ----- Document grid (`w:docGrid`) parsing tests (issue #518) -----

/// A one-paragraph document whose section carries `<w:docGrid w:linePitch="360"
/// {type_attribute}>`, written as raw XML because docx-rs's builder cannot
/// place a `w:docGrid` on the body section.
fn build_docx_with_doc_grid(type_attribute: &str) -> Vec<u8> {
    use std::io::Write;
    use zip::ZipWriter;
    use zip::write::FileOptions;

    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
    )
    .unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
    )
    .unwrap();

    zip.start_file("word/document.xml", opts).unwrap();
    let document_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t xml:space="preserve">본문 한 줄</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:docGrid w:linePitch="360"{type_attribute}/>
    </w:sectPr>
  </w:body>
</w:document>"#
    );
    zip.write_all(document_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn parse_flow_page(data: &[u8]) -> FlowPage {
    let (doc, _warnings) = DocxParser
        .parse(data, &ConvertOptions::default())
        .expect("document parses");
    match &doc.pages[0] {
        Page::Flow(flow) => flow.clone(),
        _ => panic!("Expected FlowPage"),
    }
}

#[test]
fn doc_grid_without_a_type_declares_a_pitch_that_does_not_snap() {
    // Word writes a bare `<w:docGrid w:linePitch="360"/>` into ordinary Korean
    // documents. `w:type` then takes its default value `default`, which is
    // ECMA-376's name for *no* grid, and Word lays the file out with none —
    // every Korean fixture in the business corpus is like this and none of
    // their line advances is a multiple of 18pt (issue #518).
    let page = parse_flow_page(&build_docx_with_doc_grid(""));

    assert_eq!(
        page.line_grid_pitch,
        Some(18.0),
        "the declared pitch is still read: it marks an East Asian edition"
    );
    assert!(
        !page.line_grid_snaps_lines,
        "a `default` grid must not snap lines to that pitch"
    );
}

#[test]
fn doc_grid_typed_lines_snaps_lines_to_the_pitch() {
    // Triangulation: the author turning the grid on is what makes it real.
    for grid_type in ["lines", "linesAndChars", "snapToChars"] {
        let page = parse_flow_page(&build_docx_with_doc_grid(&format!(
            r#" w:type="{grid_type}""#
        )));

        assert_eq!(page.line_grid_pitch, Some(18.0));
        assert!(
            page.line_grid_snaps_lines,
            "w:type=\"{grid_type}\" snaps lines to the grid"
        );
    }
}

#[test]
fn a_section_without_a_doc_grid_has_no_pitch_at_all() {
    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body")));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();

    let page = parse_flow_page(&cursor.into_inner());

    assert_eq!(page.line_grid_pitch, None);
    assert!(!page.line_grid_snaps_lines);
}

// ----- Body paragraph `w:pBdr w:space` parsing (issue #520) -----

/// A one-paragraph document whose only paragraph carries a bottom rule with
/// the given `w:space`, in points.
fn build_docx_with_paragraph_rule(space: Option<usize>) -> Vec<u8> {
    let mut border = docx_rs::ParagraphBorder::new(docx_rs::ParagraphBorderPosition::Bottom)
        .val(docx_rs::BorderType::Double)
        .size(8)
        .color("000000");
    if let Some(space) = space {
        border = border.space(space);
    }
    let mut paragraph =
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Letterhead"));
    paragraph.property = paragraph
        .property
        .set_borders(docx_rs::ParagraphBorders::with_empty().set(border));

    let docx = docx_rs::Docx::new().add_paragraph(paragraph);
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

fn first_body_paragraph_style(data: &[u8]) -> ParagraphStyle {
    let (doc, _warnings) = DocxParser
        .parse(data, &ConvertOptions::default())
        .expect("document parses");
    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };
    match &flow.content[0] {
        Block::Paragraph(paragraph) => paragraph.style.clone(),
        other => panic!("Expected a paragraph, got {other:?}"),
    }
}

#[test]
fn body_paragraph_rule_carries_its_declared_space() {
    // The gap between a paragraph's text and its rule is the paragraph's own
    // `w:space`, in points. Substituting a fixed 4pt displaced everything
    // below a bordered paragraph by the difference (issue #520).
    let style = first_body_paragraph_style(&build_docx_with_paragraph_rule(Some(8)));

    let space = style.border_space.expect("w:space parsed");
    assert_eq!(space.bottom, 8.0);
    assert_eq!((space.top, space.left, space.right), (0.0, 0.0, 0.0));
}

#[test]
fn a_rule_without_w_space_yields_no_gap() {
    // Triangulation: the attribute's own default is 0, so an omitted `w:space`
    // must not resurrect a house value.
    let style = first_body_paragraph_style(&build_docx_with_paragraph_rule(None));

    assert!(style.border.is_some(), "the rule itself is still parsed");
    assert_eq!(style.border_space.map(|space| space.bottom), Some(0.0));
}

#[test]
fn a_paragraph_without_a_rule_has_no_border_space() {
    let docx = docx_rs::Docx::new()
        .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body")));
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();

    let style = first_body_paragraph_style(&cursor.into_inner());

    assert!(style.border.is_none());
    assert!(style.border_space.is_none());
}

// ----- Word's East Asian/Latin auto space (issue #521) -----

/// The in-text marker the parser places at such a boundary. Duplicated from
/// the parser so the test pins the wire format rather than the constant.
const AUTO_SPACE_MARKER: char = '\u{E001}';

/// A one-paragraph document whose only run holds `text`, optionally aligned.
fn build_docx_with_korean_text_aligned(
    text: &str,
    alignment: Option<docx_rs::AlignmentType>,
) -> Vec<u8> {
    let mut paragraph = docx_rs::Paragraph::new().add_run(
        docx_rs::Run::new()
            .add_text(text)
            .fonts(docx_rs::RunFonts::new().east_asia("Malgun Gothic"))
            .size(21),
    );
    if let Some(alignment) = alignment {
        paragraph = paragraph.align(alignment);
    }
    let docx = docx_rs::Docx::new().add_paragraph(paragraph);
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// The same document, left to Word's default alignment.
fn build_docx_with_korean_text(text: &str) -> Vec<u8> {
    build_docx_with_korean_text_aligned(text, None)
}

fn first_paragraph_text(data: &[u8]) -> String {
    let (doc, _warnings) = DocxParser
        .parse(data, &ConvertOptions::default())
        .expect("document parses");
    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };
    match &flow.content[0] {
        Block::Paragraph(paragraph) => paragraph.runs.iter().map(|run| run.text.as_str()).collect(),
        other => panic!("Expected a paragraph, got {other:?}"),
    }
}

#[test]
fn a_boundary_between_east_asian_text_and_a_number_carries_the_auto_space() {
    // Word inserts a quarter em where East Asian text meets a Latin letter or
    // digit with no literal space between, on both sides of the island. A
    // native export measures 2.625pt at 10.5pt and 2.375pt at 9.5pt, and our
    // output was that much narrower at every such boundary (issue #521).
    let text = first_paragraph_text(&build_docx_with_korean_text("2026년 제3자"));

    assert_eq!(
        text,
        format!("2026{AUTO_SPACE_MARKER}년 제{AUTO_SPACE_MARKER}3{AUTO_SPACE_MARKER}자"),
        "both sides of a digit island widen, and only boundaries without a \
         literal space do"
    );
}

#[test]
fn a_boundary_that_already_has_a_space_gets_nothing() {
    // Triangulation: Word adds nothing where the author already typed a space,
    // which is why `은 2026` measures the same in the GT as in our output.
    let text = first_paragraph_text(&build_docx_with_korean_text("유효기간은 2026"));

    assert!(
        !text.contains(AUTO_SPACE_MARKER),
        "a literal space already separates the two scripts: {text:?}"
    );
}

#[test]
fn an_aligned_paragraph_widens_exactly_like_an_unaligned_one() {
    // The #1053 one-factor probe patched only `w:jc` in a Normal-defining
    // package and exported each variant through native Word: left, centred,
    // justified and right all measure +2.588pt at every boundary of an
    // unstretched line. Alignment is not part of the predicate.
    let expected = format!("2026{AUTO_SPACE_MARKER}년 제{AUTO_SPACE_MARKER}3{AUTO_SPACE_MARKER}자");

    for alignment in [
        docx_rs::AlignmentType::Both,
        docx_rs::AlignmentType::Center,
        docx_rs::AlignmentType::Right,
    ] {
        let text = first_paragraph_text(&build_docx_with_korean_text_aligned(
            "2026년 제3자",
            Some(alignment),
        ));

        assert_eq!(
            text, expected,
            "{alignment:?} takes the same quarter em as an unaligned paragraph"
        );
    }
}

#[test]
fn an_aligned_paragraph_stays_flush_without_a_defined_default_style() {
    // Triangulation on the other factor: alignment does not *grant* the space
    // either. In a corpus-shaped package — no default style defined — Word
    // draws a centred or justified line flush, which is why 02_contract_ko's
    // centred date line (issue #728) and its justified body are flush in the
    // GT. Only the style resolution decides (issue #732).
    for alignment in ["center", "both"] {
        let body = format!(
            r#"<w:p><w:pPr><w:jc w:val="{alignment}"/></w:pPr>{}</w:p>"#,
            korean_run_xml("2026년 제3자")
        );
        let text = first_paragraph_text(&build_docx_with_raw_styles(CORPUS_SHAPED_STYLES, &body));

        assert_eq!(
            text, "2026년 제3자",
            "a bare {alignment} paragraph keeps the boundaries the author typed"
        );
    }
}

#[test]
fn latin_only_and_east_asian_only_text_are_untouched() {
    // Triangulation on both sides of the predicate: the rule needs one of each
    // script, so neither a pure Latin run nor a pure Korean one may widen.
    for text in ["Version 2026 release 3", "계약서를 작성하여 보관한다"] {
        let parsed = first_paragraph_text(&build_docx_with_korean_text(text));
        assert!(
            !parsed.contains(AUTO_SPACE_MARKER),
            "single-script text needs no auto space: {parsed:?}"
        );
    }
}

#[test]
fn cjk_punctuation_is_not_a_boundary() {
    // `is_east_asian_text` is deliberately narrower than the renderer's
    // `is_cjk_like`: CJK punctuation and the fullwidth forms are already
    // full-width, and Word adds nothing beside them.
    let text = first_paragraph_text(&build_docx_with_korean_text("、2026"));

    assert!(
        !text.contains(AUTO_SPACE_MARKER),
        "an ideographic comma is already full-width: {text:?}"
    );
}

// ----- The trigger is a defined paragraph style (issues #627, #732) -----

/// A minimal package whose `word/styles.xml` is under the test's control,
/// written as raw XML because docx-rs always writes a `Normal` definition
/// into the styles part it builds — and whether the default paragraph style
/// is *defined at all* is exactly the factor these tests vary (issue #732).
///
/// Shared with the style tests, which vary `w:docDefaults/w:rPrDefault/w:rFonts`
/// the same way (issue #1196).
pub(super) fn build_docx_with_raw_styles(styles_xml: &str, body_xml: &str) -> Vec<u8> {
    use std::io::Write;
    use zip::ZipWriter;
    use zip::write::FileOptions;

    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#,
    )
    .unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
    )
    .unwrap();

    zip.start_file("word/_rels/document.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    )
    .unwrap();

    zip.start_file("word/styles.xml", opts).unwrap();
    zip.write_all(styles_xml.as_bytes()).unwrap();

    zip.start_file("word/document.xml", opts).unwrap();
    let document_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body_xml}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>"#
    );
    zip.write_all(document_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

/// The styles part every Korean business mock ships: document defaults plus a
/// `ListParagraph` definition, and — decisively — no `Normal` (issue #732).
const CORPUS_SHAPED_STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Malgun Gothic" w:cs="Malgun Gothic" w:eastAsia="Malgun Gothic" w:hAnsi="Malgun Gothic"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>
</w:styles>"#;

/// The same styles part with an explicit default paragraph style added — the
/// one factor #521's probe package differed by. Deliberately not named
/// `Normal`, so the test isolates the `w:default="1"` arm of the scan.
const DEFAULT_STYLE_DEFINING_STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Malgun Gothic" w:cs="Malgun Gothic" w:eastAsia="Malgun Gothic" w:hAnsi="Malgun Gothic"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Standard" w:default="1"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/></w:style>
</w:styles>"#;

/// A run in the corpus mocks' shape: explicit Malgun Gothic, 9.5pt.
fn korean_run_xml(text: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rFonts w:ascii="Malgun Gothic" w:cs="Malgun Gothic" w:eastAsia="Malgun Gothic" w:hAnsi="Malgun Gothic"/><w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#
    )
}

/// `text` inside `nesting` levels of single-cell table, as raw body XML.
fn nested_cell_body_xml(text: &str, nesting: usize) -> String {
    let mut body = format!("<w:p>{}</w:p>", korean_run_xml(text));
    for _ in 0..nesting {
        body = format!(
            "<w:tbl><w:tblPr><w:tblW w:type=\"dxa\" w:w=\"9026\"/></w:tblPr>\
             <w:tblGrid><w:gridCol w:w=\"9026\"/></w:tblGrid>\
             <w:tr><w:tc><w:tcPr><w:tcW w:type=\"dxa\" w:w=\"9026\"/></w:tcPr>{body}</w:tc></w:tr></w:tbl>"
        );
    }
    // A table may not end the body, so Word always writes a paragraph after it.
    body.push_str("<w:p/>");
    body
}

#[test]
fn a_bare_paragraph_stays_flush_without_a_defined_default_style() {
    // The plain-body rule of issue #732, settled by a one-factor probe: in a
    // package that defines no default paragraph style — the shape of every
    // Korean business mock — native Word draws every digit-Hangul boundary of
    // a bare paragraph flush (06_official_letter_ko `1→부` is -0.06pt), and
    // adding only a default-style definition to the same bytes flips every one
    // of them to +0.25em. Word's own built-in Korean Normal carries the
    // suppression; a defined style replaces it.
    let text = first_paragraph_text(&build_docx_with_raw_styles(
        CORPUS_SHAPED_STYLES,
        &format!("<w:p>{}</w:p>", korean_run_xml("체크리스트 1부.")),
    ));

    assert_eq!(
        text, "체크리스트 1부.",
        "a bare paragraph keeps the boundaries the author typed"
    );
}

#[test]
fn a_bare_paragraph_widens_when_a_default_style_is_defined() {
    // The other side of the same factor: #521's probe package defined its
    // `Normal` explicitly, and native Word gave its bare paragraphs the full
    // 0.25em at every boundary — 2.62pt at 10.5pt, 2.37pt at 9.5pt. Only the
    // styles part differs from the fixture above.
    let text = first_paragraph_text(&build_docx_with_raw_styles(
        DEFAULT_STYLE_DEFINING_STYLES,
        &format!("<w:p>{}</w:p>", korean_run_xml("체크리스트 1부.")),
    ));

    assert_eq!(
        text,
        format!("체크리스트 1{AUTO_SPACE_MARKER}부."),
        "a defined default style reactivates the auto space"
    );
}

#[test]
fn a_list_styled_paragraph_widens_without_numbering() {
    // #521's first survey read the corpus split as a `w:numPr` correlation;
    // the probe shows the style is the real trigger: a `ListParagraph`-styled
    // paragraph with no numbering at all takes the full 0.25em in the same
    // package whose bare paragraphs are flush (case H, issue #732). This is
    // why 02/03's list items widen while 06's plain body does not.
    let body = format!(
        r#"<w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>{}</w:p>"#,
        korean_run_xml("2026년 8월")
    );
    let text = first_paragraph_text(&build_docx_with_raw_styles(CORPUS_SHAPED_STYLES, &body));

    assert_eq!(
        text,
        format!("2026{AUTO_SPACE_MARKER}년 8{AUTO_SPACE_MARKER}월"),
        "a defined, referenced style reactivates the auto space"
    );
}

/// The same run `build_docx_with_korean_text` produces, wrapped in `nesting`
/// levels of single-cell table so a cell inside a cell can be exercised too.
fn build_docx_with_korean_text_in_table_cell(text: &str, nesting: usize) -> Vec<u8> {
    let paragraph = docx_rs::Paragraph::new().add_run(
        docx_rs::Run::new()
            .add_text(text)
            .fonts(docx_rs::RunFonts::new().east_asia("Malgun Gothic"))
            .size(21),
    );
    let mut table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
        docx_rs::TableCell::new().add_paragraph(paragraph),
    ])])
    .set_grid(vec![4000]);
    for _ in 1..nesting {
        table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_table(table),
        ])])
        .set_grid(vec![4000]);
    }
    let docx = docx_rs::Docx::new().add_table(table);
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// The run text of the first paragraph in the innermost cell, descending
/// through however many nested tables stand in the way.
fn innermost_cell_paragraph_text(data: &[u8]) -> String {
    let (doc, _warnings) = DocxParser
        .parse(data, &ConvertOptions::default())
        .expect("document parses");
    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        other => panic!("Expected FlowPage, got {other:?}"),
    };
    fn descend(blocks: &[Block]) -> Option<String> {
        blocks.iter().find_map(|block| match block {
            Block::Table(table) => descend(&table.rows[0].cells[0].content),
            Block::Paragraph(paragraph) => {
                Some(paragraph.runs.iter().map(|run| run.text.as_str()).collect())
            }
            _ => None,
        })
    }
    descend(&flow.content).expect("a paragraph inside the innermost cell")
}

#[test]
fn a_table_cell_gets_no_auto_space_at_a_digit_hangul_boundary() {
    // In 10_research_report_ko's month column every digit-Hangul boundary is
    // flush (-0.056pt) in the native export while ours opened 2.366pt, taking
    // the cell text from Word's 48.60pt to 53.25pt (+9.6%) and rendering
    // `2024년 1월` as `2024 년 1 월` (issue #627). #627 read this as a cell
    // rule; the #732 probe shows it is the same style rule as the body: those
    // cell paragraphs are bare, and the corpus defines no default style.
    let text = innermost_cell_paragraph_text(&build_docx_with_raw_styles(
        CORPUS_SHAPED_STYLES,
        &nested_cell_body_xml("2024년 1월", 1),
    ));

    assert_eq!(
        text, "2024년 1월",
        "cell text keeps the boundaries the author typed"
    );
}

#[test]
fn a_body_paragraph_is_left_exactly_as_it_was() {
    // docx-rs writes a `Normal` definition into every styles part it builds,
    // so this document is in #521's-probe territory, not the corpus's: its
    // probe defined `Normal` too, and native Word widened its bare paragraphs
    // at every boundary (issue #732). The assertion is unchanged from when it
    // pinned today's emission — what changed is that the emission is now known
    // to be what Word does to this document.
    let text = first_paragraph_text(&build_docx_with_korean_text("2024년 1월"));

    assert_eq!(
        text,
        format!("2024{AUTO_SPACE_MARKER}년 1{AUTO_SPACE_MARKER}월"),
        "a document whose builder defines Normal keeps the auto space"
    );
}

#[test]
fn a_table_cell_widens_when_a_default_style_is_defined() {
    // #521's probe put the same sentence in a table cell of its
    // Normal-defining package and native Word widened every boundary exactly
    // as in the body (cases I/J: 2.62pt at 10.5pt, 2.37pt at 9.5pt). The cell
    // is not a suppressor — the undefined default style is, and this docx-rs
    // package defines one (issue #732).
    let text =
        innermost_cell_paragraph_text(&build_docx_with_korean_text_in_table_cell("2024년 1월", 1));

    assert_eq!(
        text,
        format!("2024{AUTO_SPACE_MARKER}년 1{AUTO_SPACE_MARKER}월"),
        "a defined default style reactivates the auto space in cells too"
    );
}

#[test]
fn a_numbered_list_paragraph_keeps_the_auto_space() {
    // The corpus GT's positive case: the `w:pStyle="ListParagraph"` + `w:numPr`
    // paragraphs of 02 and 03 widen (8.41/8.40pt, matched by us today). The
    // #732 probe shows the referenced style — not the numbering — carries
    // this, so a numbered item must keep widening under the style predicate.
    let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("%1."),
        docx_rs::LevelJc::new("left"),
    ));
    let data = build_docx_with_numbering(
        vec![abstract_num],
        vec![docx_rs::Numbering::new(1, 0)],
        vec![
            docx_rs::Paragraph::new()
                .add_run(
                    docx_rs::Run::new()
                        .add_text("2024년 1월")
                        .fonts(docx_rs::RunFonts::new().east_asia("Malgun Gothic"))
                        .size(21),
                )
                .style("ListParagraph")
                .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        ],
    );

    let (doc, _warnings) = DocxParser
        .parse(&data, &ConvertOptions::default())
        .expect("document parses");
    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        other => panic!("Expected FlowPage, got {other:?}"),
    };
    let text: String = flow
        .content
        .iter()
        .find_map(|block| match block {
            Block::List(list) => Some(
                list.items[0]
                    .content
                    .iter()
                    .flat_map(|paragraph| paragraph.runs.iter().map(|run| run.text.as_str()))
                    .collect::<String>(),
            ),
            _ => None,
        })
        .expect("a list block");

    assert_eq!(
        text,
        format!("2024{AUTO_SPACE_MARKER}년 1{AUTO_SPACE_MARKER}월"),
        "a numbered list item is body text and still widens"
    );
}

#[test]
fn a_cell_inside_a_cell_is_still_a_cell() {
    // Triangulation on depth: the style rule reaches a nested table's bare
    // cell paragraphs the same way it reaches the outer table's.
    let text = innermost_cell_paragraph_text(&build_docx_with_raw_styles(
        CORPUS_SHAPED_STYLES,
        &nested_cell_body_xml("2024년 1월", 2),
    ));

    assert_eq!(
        text, "2024년 1월",
        "a nested cell's bare paragraph stays flush too"
    );
}

// ----- `w:titlePg` selects the first-page stories (issue #846) -----

/// Every run's text in a header or footer story, concatenated.
fn header_footer_text(story: &crate::ir::HeaderFooter) -> String {
    story
        .paragraphs
        .iter()
        .flat_map(|paragraph| paragraph.elements.iter())
        .filter_map(|element| match element {
            crate::ir::HFInline::Run(run) => Some(run.text.as_str()),
            _ => None,
        })
        .collect()
}

fn build_docx_with_title_page(title_pg: bool) -> Vec<u8> {
    let default_header = docx_rs::Header::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Every other page")),
    );
    let first_header = docx_rs::Header::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cover banner")),
    );
    let docx = docx_rs::Docx::new()
        .header(default_header)
        .first_header(first_header)
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
        );
    let docx = if title_pg { docx.title_pg() } else { docx };
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn title_pg_keeps_the_first_page_header_apart_from_the_rest() {
    let page = parse_flow_page(&build_docx_with_title_page(true));

    let first = page
        .first_header
        .as_ref()
        .expect("titlePg must resolve a first-page header");
    assert!(
        header_footer_text(first).contains("Cover banner"),
        "the first page takes the `first` story, got {:?}",
        header_footer_text(first)
    );
    let rest = page
        .header
        .as_ref()
        .expect("the remaining pages keep the default header");
    assert!(
        header_footer_text(rest).contains("Every other page"),
        "pages after the first keep the default story, got {:?}",
        header_footer_text(rest)
    );
}

#[test]
fn a_first_header_without_title_pg_is_not_used() {
    // Triangulation: Word only honours the `first` story when `w:titlePg` asks
    // for it, so its mere presence must not split the section.
    //
    // Driven off `SectionProperty` rather than a built document because
    // `docx_rs`'s `first_header` sets `title_pg` for you — its sibling is
    // named `first_header_without_title_pg` precisely because the two are
    // separable in the format even though that builder couples them.
    let story = || {
        docx_rs::Header::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cover banner")),
        )
    };
    let assets = super::sections::HeaderFooterAssets::default();
    let style_map = super::StyleMap::new();
    let styles = super::sections::HeaderFooterStyleContext {
        style_map: &style_map,
        paragraph_property_defaults_are_declared: false,
    };

    let declared = docx_rs::SectionProperty::new().first_header(story(), "rId9");
    assert!(
        declared.title_pg,
        "the builder is expected to set title_pg here; if this fails the \
         coupling changed and the test below proves nothing"
    );
    assert!(
        super::sections::extract_docx_first_header(&declared, &assets, styles).is_some(),
        "titlePg with a first story resolves it"
    );

    let undeclared = docx_rs::SectionProperty::new().first_header_without_title_pg(story(), "rId9");
    assert!(!undeclared.title_pg);
    assert!(
        super::sections::extract_docx_first_header(&undeclared, &assets, styles).is_none(),
        "a first story without titlePg is not a first-page story"
    );
}

#[test]
fn a_title_page_header_is_emitted_as_a_per_page_choice() {
    // The two stories become one `header:` value that asks which page it is
    // on, rather than the section committing to a single story (issue #846).
    let data = build_docx_with_title_page(true);
    let (document, _warnings) = DocxParser
        .parse(&data, &ConvertOptions::default())
        .expect("document parses");
    let source = crate::internal::generate_typst(&document)
        .expect("document generates")
        .source;

    assert!(
        source.contains("header: context { if here().page() =="),
        "the header must choose per page, got:\n{source}"
    );
    assert!(
        source.contains("query(<o2p-sec-0>).first().location().page()"),
        "the choice resolves the section's own first page, got:\n{source}"
    );
    assert!(
        source.contains("#metadata(none) <o2p-sec-0>"),
        "the section must label its first page, got:\n{source}"
    );
    let cover = source.find("Cover banner").expect("first story present");
    let rest = source
        .find("Every other page")
        .expect("default story present");
    assert!(
        cover < rest,
        "the first-page story is the `if` branch and the default the `else`"
    );
}

#[test]
fn a_section_with_only_a_first_page_header_still_emits_it() {
    // A `w:titlePg` section may declare just the `first` story, meaning later
    // pages carry no header at all. The page-setup shortcut used to ask only
    // about the default stories and would have dropped this one (issue #846).
    let data = build_docx_with_title_page(true);
    let (mut document, _warnings) = DocxParser
        .parse(&data, &ConvertOptions::default())
        .expect("document parses");
    if let Some(crate::ir::Page::Flow(page)) = document.pages.first_mut() {
        page.header = None;
        page.footer = None;
    }
    let source = crate::internal::generate_typst(&document)
        .expect("document generates")
        .source;
    assert!(
        source.contains("Cover only") || source.contains("Cover banner"),
        "a first-only header must still reach the page setup, got:\n{source}"
    );
}

// ----- Hangul wrap follows the same defined-style trigger (issue #833) -----

/// A pure-Hangul sentence — no digit/Latin boundary, so the auto-space pass
/// leaves it byte-identical and the assertions below see only the wrap frames.
const HANGUL_WRAP_SENTENCE: &str = "본 계약은 갑과 을이";

/// Framing puts each multi-syllable eojeol alone inside a bracket pair
/// (`…#box(…)[#text(…)[계약은]]…`), where the unframed sentence emits as one
/// span (`[본 계약은 갑과 을이]`) — so this substring appears iff framed.
const FRAMED_EOJEOL: &str = "[계약은]";

/// End-to-end: package bytes through the parser and the Typst generator, so
/// the assertion sees the frames a paragraph would actually wrap with.
fn typst_source_for(styles_xml: &str, body_xml: &str) -> String {
    let (doc, _warnings) = DocxParser
        .parse(
            &build_docx_with_raw_styles(styles_xml, body_xml),
            &ConvertOptions::default(),
        )
        .expect("document parses");
    crate::internal::generate_typst(&doc)
        .expect("document generates")
        .source
}

#[test]
fn a_bare_hangul_paragraph_gets_no_eojeol_frames_without_a_defined_default_style() {
    // The #833 probe series: the report's bare 9pt note breaks `표시되어야`
    // after `표` in native Word (GT line 1 ends at 523.68pt), and so does the
    // same text at 10.5pt, without italic, and without its final stop — while
    // re-exporting the same bytes plus only a default-style definition keeps
    // the eojeol whole with `표` declined at 524.1pt of a 524.45pt measure.
    // Word's built-in Korean Normal breaks Hangul at character level; every
    // paragraph #626 measured as eojeol-whole is `ListParagraph`-styled.
    let source = typst_source_for(
        CORPUS_SHAPED_STYLES,
        &format!("<w:p>{}</w:p>", korean_run_xml(HANGUL_WRAP_SENTENCE)),
    );

    assert!(
        source.contains(HANGUL_WRAP_SENTENCE),
        "a bare paragraph keeps syllable-level break opportunities: {source}"
    );
    assert!(
        !source.contains(FRAMED_EOJEOL),
        "no eojeol frame may reach a bare paragraph: {source}"
    );
}

#[test]
fn a_style_referencing_hangul_paragraph_keeps_its_eojeol_frames() {
    // Probe F2: `w:pStyle="ListParagraph"` alone — no numbering — keeps
    // `정함을` whole with 24.85pt to spare, in the same package whose bare
    // control breaks mid-word. Probe F7 measures the same for `Heading6`, so
    // the trigger is any resolvable style, not the list machinery.
    let body = format!(
        r#"<w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>{}</w:p>"#,
        korean_run_xml(HANGUL_WRAP_SENTENCE)
    );
    let source = typst_source_for(CORPUS_SHAPED_STYLES, &body);

    assert!(
        source.contains(FRAMED_EOJEOL),
        "a defined, referenced style keeps each eojeol whole: {source}"
    );
}

#[test]
fn a_bare_hangul_paragraph_keeps_its_eojeol_frames_when_a_default_style_is_defined() {
    // Probe G1: adding only a default-style definition to the report package
    // flips its bare note paragraph from breaking after `표` to keeping
    // `표시되어야` whole — the leg every real Word-authored document is on,
    // since Word always writes a `Normal` definition.
    let source = typst_source_for(
        DEFAULT_STYLE_DEFINING_STYLES,
        &format!("<w:p>{}</w:p>", korean_run_xml(HANGUL_WRAP_SENTENCE)),
    );

    assert!(
        source.contains(FRAMED_EOJEOL),
        "a defined default style keeps each eojeol whole: {source}"
    );
}

#[test]
fn an_unresolvable_pstyle_gets_no_eojeol_frames() {
    // Probe F8: a `w:pStyle` naming a style the document never defines wraps
    // exactly like the bare control — mid-word — so resolution, not the mere
    // presence of the reference, is what replaces the built-in Korean Normal.
    let body = format!(
        r#"<w:p><w:pPr><w:pStyle w:val="NoSuchStyle"/></w:pPr>{}</w:p>"#,
        korean_run_xml(HANGUL_WRAP_SENTENCE)
    );
    let source = typst_source_for(CORPUS_SHAPED_STYLES, &body);

    assert!(
        !source.contains(FRAMED_EOJEOL),
        "an unresolvable style reference is the same as no style: {source}"
    );
}

#[test]
fn a_bare_cell_hangul_paragraph_gets_no_eojeol_frames_without_a_defined_default_style() {
    // Probe H2: a bare cell paragraph in the no-default-style package breaks
    // `목적으로` after `목` at the cell's own measure — the cell follows the
    // same style rule as the body, as it does for the auto space (#627, #732).
    let source = typst_source_for(
        CORPUS_SHAPED_STYLES,
        &nested_cell_body_xml(HANGUL_WRAP_SENTENCE, 1),
    );

    assert!(
        source.contains(HANGUL_WRAP_SENTENCE),
        "a bare cell paragraph keeps syllable-level break opportunities: {source}"
    );
    assert!(
        !source.contains(FRAMED_EOJEOL),
        "no eojeol frame may reach a bare cell paragraph: {source}"
    );
}

#[test]
fn an_explicit_word_wrap_keeps_eojeol_frames_on_a_bare_paragraph() {
    // Direct formatting outranks the style chain it overrides (issue #730's
    // `w:val="0"` is checked first for the same reason), so an explicit
    // `w:wordWrap w:val="1"` restores word-level wrapping even where the
    // built-in Korean Normal would break characters.
    let body = format!(
        r#"<w:p><w:pPr><w:wordWrap w:val="1"/></w:pPr>{}</w:p>"#,
        korean_run_xml(HANGUL_WRAP_SENTENCE)
    );
    let source = typst_source_for(CORPUS_SHAPED_STYLES, &body);

    assert!(
        source.contains(FRAMED_EOJEOL),
        "an explicit word-level wrap request keeps the frames: {source}"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_header_tab_before_a_centre_stop_centres_its_text_on_that_stop() {
    // Word's Header style declares a centre and a right stop, and a centred
    // running head is typed as `<tab>Title`, which Word saves as one run. The
    // tab advances to the centre stop, the first one past the pen. Header
    // tabs were matched against running-head shapes instead, and a lone tab
    // with a right stop last went to the right margin (issue #1821). The same
    // paragraph in the body is the reference: both must land on one x.
    let centred_head = || {
        docx_rs::Paragraph::new()
            .add_tab(
                docx_rs::Tab::new()
                    .val(docx_rs::TabValueType::Center)
                    .pos(4153),
            )
            .add_tab(
                docx_rs::Tab::new()
                    .val(docx_rs::TabValueType::Right)
                    .pos(8306),
            )
            .add_run(docx_rs::Run::new().add_tab().add_text("{Title}"))
    };
    let docx = docx_rs::Docx::new()
        .header(docx_rs::Header::new().add_paragraph(centred_head()))
        .add_paragraph(centred_head());
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data: Vec<u8> = cursor.into_inner();

    let left_margin: f64 = parse_flow_page(&data).margins.left;
    let (document, _warnings) = DocxParser
        .parse(&data, &ConvertOptions::default())
        .expect("document parses");
    let source: String = crate::internal::generate_typst(&document)
        .expect("document generates")
        .source;
    let mut placed: Vec<(f64, f64)> = crate::render::pdf::compiled_text_runs(&source, 0)
        .unwrap()
        .into_iter()
        .filter(|run| run.text == "{Title}")
        .map(|run| (run.baseline_pt, run.left_pt))
        .collect();
    placed.sort_by(|a, b| a.0.total_cmp(&b.0));
    let [(_, header_x), (_, body_x)] = placed[..] else {
        panic!("expected the title in the header and in the body, got {placed:?}");
    };

    let centre_stop: f64 = left_margin + 4153.0 / 20.0;
    assert!(
        (centre_stop - 40.0..centre_stop).contains(&body_x),
        "the body centres the title on the {centre_stop}pt stop: {body_x}pt"
    );
    assert!(
        (header_x - body_x).abs() < 0.01,
        "the header title starts at {header_x}pt, the body's at {body_x}pt"
    );
}
