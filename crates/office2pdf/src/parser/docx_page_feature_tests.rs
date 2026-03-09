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
            .any(|element| matches!(element, crate::ir::HFInline::PageNumber))
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
            .any(|element| matches!(element, crate::ir::HFInline::TotalPages))
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
