use super::*;
use crate::ir::HeaderFooterParagraph;
use crate::render::pdf::{PlacedTextRun, compiled_text_runs_with_fonts_and_images};

fn footer(lines: usize, distance: f64, size: f64) -> HeaderFooter {
    HeaderFooter {
        shapes: vec![],
        distance_from_edge: Some(distance),
        sheet_print_scale: None,
        paragraphs: (0..lines)
            .map(|index| HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: format!("Footer {index}"),
                    style: TextStyle {
                        font_family: Some("Noto Sans".into()),
                        font_size: Some(size),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                    inline_box: None,
                })],
                border: None,
                border_space: None,
                sheet_section_is_rich: false,
                frame: None,
            })
            .collect(),
    }
}

fn probe(bottom: f64, footer: Option<HeaderFooter>) -> Document {
    let content = (0..60)
        .map(|index| {
            let Block::Paragraph(mut paragraph) = make_paragraph(&format!("Body {index:02}"))
            else {
                unreachable!()
            };
            paragraph.style.line_spacing = Some(LineSpacing::Exact(12.0));
            paragraph.style.space_before = Some(0.0);
            paragraph.style.space_after = Some(0.0);
            paragraph.runs[0].style.font_family = Some("Noto Sans".into());
            paragraph.runs[0].style.font_size = Some(10.0);
            Block::Paragraph(paragraph)
        })
        .collect();
    let Page::Flow(mut page) = make_flow_page(content) else {
        unreachable!()
    };
    page.size = PageSize {
        width: 260.0,
        height: 300.0,
    };
    page.margins = Margins {
        top: 24.0,
        bottom,
        left: 24.0,
        right: 24.0,
    };
    page.footer = footer;
    make_doc(vec![Page::Flow(page)])
}

fn runs(doc: &Document, page: usize) -> Vec<PlacedTextRun> {
    let fonts = crate::bundled_fonts::noto_sans_fonts();
    let context = crate::render::font_context::resolve_font_search_context_from_fonts(fonts);
    let output = generate_typst_with_options_and_font_context(
        doc,
        &ConvertOptions::default(),
        Some(&context),
    )
    .expect("probe generates");
    compiled_text_runs_with_fonts_and_images(&output.source, &output.images, page, fonts)
        .expect("probe compiles")
}

fn body(runs: &[PlacedTextRun]) -> Vec<String> {
    runs.iter()
        .filter(|run| run.text.starts_with("Body "))
        .map(|run| run.text.clone())
        .collect()
}

#[test]
fn footer_extent_reserves_body_space_for_nonnegative_bottom_margins() {
    for (lines, distance, size) in [(3, 36.0, 12.0), (1, 75.0, 18.0)] {
        let doc = probe(36.0, Some(footer(lines, distance, size)));
        for page in [0, 1] {
            let placed = runs(&doc, page);
            let last_body = placed
                .iter()
                .filter(|run| run.text.starts_with("Body "))
                .map(|run| run.baseline_pt)
                .reduce(f64::max)
                .unwrap();
            let first_footer = placed
                .iter()
                .filter(|run| run.text.starts_with("Footer "))
                .map(|run| run.baseline_pt)
                .reduce(f64::min)
                .unwrap();
            assert!(
                last_body + size * 0.8 < first_footer,
                "body must end above the footer: body={last_body}, footer={first_footer}, lines={lines}, page={page}"
            );
        }
    }
}

#[test]
fn footer_extent_keeps_the_page_edge_anchor_when_body_margin_grows() {
    let story = footer(3, 36.0, 12.0);
    let overflowing = runs(&probe(36.0, Some(story.clone())), 0);
    let roomy = runs(&probe(110.0, Some(story)), 0);
    for index in 0..3 {
        let label = format!("Footer {index}");
        let baseline = |runs: &[PlacedTextRun]| {
            runs.iter()
                .find(|run| run.text == label)
                .unwrap()
                .baseline_pt
        };
        assert!(
            (baseline(&overflowing) - baseline(&roomy)).abs() < 0.01,
            "the same footer must stay anchored to the page edge"
        );
    }
}

#[test]
fn footer_extent_negative_margin_uses_absolute_body_boundary_and_allows_overlap() {
    let fixed = runs(&probe(-36.0, Some(footer(3, 50.0, 12.0))), 0);
    let no_footer = runs(&probe(36.0, None), 0);
    assert_eq!(
        body(&fixed),
        body(&no_footer),
        "negative bottom margin keeps the absolute boundary independent of footer height"
    );
}

#[test]
fn footer_extent_first_page_reservation_does_not_shrink_later_pages() {
    let short = footer(1, 12.0, 10.0);
    let mut doc = probe(36.0, Some(short.clone()));
    let Page::Flow(page) = &mut doc.pages[0] else {
        unreachable!()
    };
    page.first_footer = Some(footer(4, 36.0, 12.0));
    let control = probe(36.0, Some(short));
    assert!(
        body(&runs(&doc, 0)).len() < body(&runs(&control, 0)).len(),
        "only the tall first-page footer reduces first-page body capacity"
    );
    assert_eq!(
        body(&runs(&doc, 1)).len(),
        body(&runs(&control, 1)).len(),
        "continuations use their own shorter footer"
    );
}

#[test]
fn footer_extent_short_first_page_can_use_more_space_than_continuations() {
    let tall = footer(4, 36.0, 12.0);
    let mut doc = probe(36.0, Some(tall.clone()));
    let Page::Flow(page) = &mut doc.pages[0] else {
        unreachable!()
    };
    page.first_footer = Some(footer(1, 12.0, 10.0));
    let control = probe(36.0, Some(tall));
    assert!(body(&runs(&doc, 0)).len() > body(&runs(&control, 0)).len());
    assert_eq!(body(&runs(&doc, 1)).len(), body(&runs(&control, 1)).len());
}

#[test]
fn footer_extent_measures_wrapped_lines_at_the_actual_section_width() {
    let mut story = footer(1, 36.0, 12.0);
    let HFInline::Run(run) = &mut story.paragraphs[0].elements[0] else {
        unreachable!()
    };
    run.text =
        "Footer words that must wrap over several lines in this narrow section of the document"
            .into();
    let mut narrow = probe(36.0, Some(story.clone()));
    let Page::Flow(page) = &mut narrow.pages[0] else {
        unreachable!()
    };
    page.size.width = 160.0;
    let placed = runs(&narrow, 0);
    let footer_lines: Vec<_> = placed
        .iter()
        .filter(|run| !run.text.starts_with("Body "))
        .collect();
    assert!(footer_lines.len() >= 3, "the footer must actually wrap");
    let first = footer_lines
        .iter()
        .map(|run| run.baseline_pt)
        .reduce(f64::min)
        .unwrap();
    let last_body = placed
        .iter()
        .filter(|run| run.text.starts_with("Body "))
        .map(|run| run.baseline_pt)
        .reduce(f64::max)
        .unwrap();
    assert!(
        last_body + 9.6 < first,
        "wrapped footer cannot overlap body"
    );
    assert!(
        body(&placed).len() < body(&runs(&probe(36.0, Some(story)), 0)).len(),
        "narrower wrapping reduces body capacity"
    );
}

#[test]
fn footer_extent_first_page_reservation_includes_footnotes() {
    fn with_note(story: HeaderFooter) -> Document {
        let mut doc = probe(36.0, Some(story));
        let Page::Flow(page) = &mut doc.pages[0] else {
            unreachable!()
        };
        let Block::Paragraph(paragraph) = &mut page.content[0] else {
            unreachable!()
        };
        paragraph.runs.push(Run {
            text: String::new(),
            style: TextStyle::default(),
            href: None,
            inline_box: None,
            footnote: Some(vec![Run {
                text: "Note text".into(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
                inline_box: None,
            }]),
        });
        doc
    }
    let control = with_note(footer(1, 12.0, 10.0));
    let mut doc = control.clone();
    let Page::Flow(page) = &mut doc.pages[0] else {
        unreachable!()
    };
    page.first_footer = Some(footer(4, 36.0, 12.0));
    let placed = runs(&doc, 0);
    let note = |runs: &[PlacedTextRun]| {
        runs.iter()
            .find(|run| run.text == "Note text")
            .unwrap()
            .baseline_pt
    };
    assert!(
        note(&placed) + 20.0 < note(&runs(&control, 0)),
        "the note region must move above the tall first-page footer"
    );
    let first_footer = placed
        .iter()
        .filter(|run| run.text.starts_with("Footer "))
        .map(|run| run.baseline_pt)
        .reduce(f64::min)
        .unwrap();
    let default_tall = runs(&with_note(footer(4, 36.0, 12.0)), 0);
    assert!(
        (note(&placed) - note(&default_tall)).abs() < 0.01,
        "the selected first-page footer must move the note region to the same boundary as an equally tall default footer"
    );
    assert!(
        note(&placed) < first_footer,
        "the note baseline precedes the first footer baseline"
    );
}

#[test]
fn footer_extent_keeps_unanchored_variant_placement_when_other_variant_is_measured() {
    let mut legacy = footer(1, 12.0, 10.0);
    legacy.distance_from_edge = None;
    let control = runs(&probe(36.0, Some(legacy.clone())), 1);
    let mut doc = probe(36.0, Some(legacy));
    let Page::Flow(page) = &mut doc.pages[0] else {
        unreachable!()
    };
    page.first_footer = Some(footer(4, 36.0, 12.0));
    let placed = runs(&doc, 1);
    let baseline = |runs: &[PlacedTextRun]| {
        runs.iter()
            .find(|run| run.text == "Footer 0")
            .unwrap()
            .baseline_pt
    };
    assert!(
        (baseline(&placed) - baseline(&control)).abs() < 0.01,
        "measuring the first variant cannot move an unanchored default footer"
    );
}

#[test]
fn footer_extent_empty_first_variant_reserves_and_paints_no_default_footer() {
    let mut doc = probe(36.0, Some(footer(4, 36.0, 12.0)));
    let Page::Flow(page) = &mut doc.pages[0] else {
        unreachable!()
    };
    page.first_footer = Some(footer(0, 36.0, 12.0));
    let first = runs(&doc, 0);
    assert!(
        !first.iter().any(|run| run.text.starts_with("Footer ")),
        "an explicit empty first footer suppresses the default story"
    );
    assert_eq!(body(&first), body(&runs(&probe(36.0, None), 0)));
    assert!(body(&runs(&doc, 1)).len() < body(&first).len());
}

#[test]
fn footer_extent_tall_first_story_keeps_every_line_at_the_selected_page_edge() {
    let tall = footer(10, 36.0, 12.0);
    let control = runs(&probe(36.0, Some(tall.clone())), 0);
    let mut doc = probe(36.0, Some(footer(1, 12.0, 10.0)));
    let Page::Flow(page) = &mut doc.pages[0] else {
        unreachable!()
    };
    page.first_footer = Some(tall);
    let actual = runs(&doc, 0);
    let story = |runs: &[PlacedTextRun]| {
        runs.iter()
            .filter(|run| run.text.starts_with("Footer "))
            .map(|run| (run.text.clone(), run.baseline_pt))
            .collect::<Vec<_>>()
    };
    let actual = story(&actual);
    let expected = story(&control);
    assert_eq!(
        actual.len(),
        expected.len(),
        "tall first story retains every line"
    );
    for ((actual_text, actual_y), (expected_text, expected_y)) in actual.iter().zip(&expected) {
        assert_eq!(actual_text, expected_text);
        assert!(
            (actual_y - expected_y).abs() < 0.01,
            "selected footer line {actual_text} must retain its physical anchor: {actual_y} vs {expected_y}"
        );
    }
}
