use super::*;

fn paragraph(text: &str) -> docx_rs::Paragraph {
    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(text))
}

fn parse(docx: docx_rs::Docx) -> Document {
    let mut bytes = Cursor::new(Vec::new());
    docx.build().pack(&mut bytes).unwrap();
    DocxParser
        .parse(&bytes.into_inner(), &ConvertOptions::default())
        .unwrap()
        .0
}

fn first_blank(header: bool, with_paragraph: bool) -> Document {
    let docx = docx_rs::Docx::new()
        .header(docx_rs::Header::new().add_paragraph(paragraph("DefaultHeader")))
        .footer(docx_rs::Footer::new().add_paragraph(paragraph("DefaultFooter")));
    let docx = if header {
        let blank = if with_paragraph {
            docx_rs::Header::new().add_paragraph(docx_rs::Paragraph::new())
        } else {
            docx_rs::Header::new()
        };
        docx.first_header(blank)
            .first_footer(docx_rs::Footer::new().add_paragraph(paragraph("DefaultFooter")))
    } else {
        let blank = if with_paragraph {
            docx_rs::Footer::new().add_paragraph(docx_rs::Paragraph::new())
        } else {
            docx_rs::Footer::new()
        };
        docx.first_footer(blank)
            .first_header(docx_rs::Header::new().add_paragraph(paragraph("DefaultHeader")))
    };
    parse(
        docx.add_paragraph(paragraph("FirstBody"))
            .add_paragraph(paragraph("SecondBody").page_break_before(true)),
    )
}

fn flow(doc: &Document, index: usize) -> &FlowPage {
    let Page::Flow(page) = &doc.pages[index] else {
        panic!("flow section")
    };
    page
}

fn painted(doc: &Document, page: usize, label: &str) -> bool {
    let output = crate::render::typst_gen::generate_typst(doc).unwrap();
    crate::render::pdf::compiled_text_runs_with_images(&output.source, &output.images, page)
        .unwrap()
        .iter()
        .any(|run| run.text.contains(label))
}

#[test]
fn empty_story_first_header_part_is_retained_as_an_explicit_override() {
    let doc = first_blank(true, false);
    let story = flow(&doc, 0)
        .first_header
        .as_ref()
        .expect("declared empty first header remains distinct from an absent reference");
    assert!(story.paragraphs.is_empty() && story.shapes.is_empty());
}

#[test]
fn empty_story_first_footer_part_is_retained_as_an_explicit_override() {
    let doc = first_blank(false, false);
    let story = flow(&doc, 0)
        .first_footer
        .as_ref()
        .expect("declared empty first footer remains distinct from an absent reference");
    assert!(story.paragraphs.is_empty() && story.shapes.is_empty());
}

#[test]
fn empty_story_first_header_suppresses_only_the_first_page_header() {
    for with_paragraph in [true, false] {
        let doc = first_blank(true, with_paragraph);
        assert!(
            !painted(&doc, 0, "DefaultHeader"),
            "explicitly blank first header cannot fall back to default"
        );
        assert!(painted(&doc, 1, "DefaultHeader"));
        for page in [0, 1] {
            assert!(painted(&doc, page, "DefaultFooter"));
        }
    }
}

#[test]
fn empty_story_first_footer_suppresses_only_the_first_page_footer() {
    for with_paragraph in [true, false] {
        let doc = first_blank(false, with_paragraph);
        assert!(
            !painted(&doc, 0, "DefaultFooter"),
            "explicitly blank first footer cannot fall back to default"
        );
        assert!(painted(&doc, 1, "DefaultFooter"));
        for page in [0, 1] {
            assert!(painted(&doc, page, "DefaultHeader"));
        }
    }
}

fn section_blank(header: bool, with_paragraph: bool) -> Document {
    let docx = docx_rs::Docx::new()
        .header(docx_rs::Header::new().add_paragraph(paragraph("EarlierHeader")))
        .footer(docx_rs::Footer::new().add_paragraph(paragraph("EarlierFooter")));
    let first = docx.document.section_property.clone();
    // Keep all parts on Docx: Section::header currently discards its pending
    // footer in docx-rs's builder. The references below still form an ordinary
    // two-section package, independently selecting each default story.
    let mut docx = if header {
        docx.first_header(if with_paragraph {
            docx_rs::Header::new().add_paragraph(docx_rs::Paragraph::new())
        } else {
            docx_rs::Header::new()
        })
    } else {
        docx.first_footer(if with_paragraph {
            docx_rs::Footer::new().add_paragraph(docx_rs::Paragraph::new())
        } else {
            docx_rs::Footer::new()
        })
    };
    let final_section = &mut docx.document.section_property;
    final_section.title_pg = false;
    if header {
        let reference = final_section.first_header_reference.take().unwrap();
        final_section.header_reference =
            Some(docx_rs::HeaderReference::new("default", &reference.id));
        final_section.footer_reference = None;
    } else {
        let reference = final_section.first_footer_reference.take().unwrap();
        final_section.footer_reference =
            Some(docx_rs::FooterReference::new("default", &reference.id));
        final_section.header_reference = None;
    }
    parse(
        docx.add_paragraph(paragraph("FirstBody").section_property(first))
            .add_paragraph(paragraph("SecondBody")),
    )
}

#[test]
fn empty_story_default_override_blocks_previous_section_but_preserves_other_story() {
    for header in [true, false] {
        for with_paragraph in [true, false] {
            let doc = section_blank(header, with_paragraph);
            let (cleared, inherited) = if header {
                ("EarlierHeader", "EarlierFooter")
            } else {
                ("EarlierFooter", "EarlierHeader")
            };
            assert!(painted(&doc, 0, cleared));
            assert!(painted(&doc, 0, inherited));
            assert!(
                !painted(&doc, 1, cleared),
                "empty default story replaces the preceding section's story"
            );
            assert!(
                painted(&doc, 1, inherited),
                "omitting the other reference retains its independent inheritance: header={header}, with_paragraph={with_paragraph}"
            );
        }
    }
}

#[test]
fn empty_story_absent_ir_slots_clear_previous_page_setup() {
    let mut doc = section_blank(true, true);
    let Page::Flow(page) = &mut doc.pages[1] else {
        unreachable!()
    };
    page.header = None;
    page.footer = None;
    assert!(!painted(&doc, 1, "EarlierHeader"));
    assert!(!painted(&doc, 1, "EarlierFooter"));
}
