use super::*;

#[test]
fn test_text_box_extraction() {
    let shape = make_text_box(0, 0, 1_000_000, 500_000, "Hello World");
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 1, "Expected 1 element");

    let blocks = text_box_blocks(&page.elements[0]);
    assert!(!blocks.is_empty(), "Expected at least one block");

    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs.len(), 1);
    assert_eq!(para.runs[0].text, "Hello World");
}

#[test]
fn test_text_box_position_and_size() {
    let x = 1_000_000i64;
    let y = 500_000i64;
    let cx = 5_000_000i64;
    let cy = 2_000_000i64;
    let shape = make_text_box(x, y, cx, cy, "Positioned");
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let elem = &page.elements[0];

    let expected_x = emu_to_pt(x);
    let expected_y = emu_to_pt(y);
    let expected_w = emu_to_pt(cx);
    let expected_h = emu_to_pt(cy);

    assert!(
        (elem.x - expected_x).abs() < 0.1,
        "Expected x ~{expected_x}, got {}",
        elem.x
    );
    assert!(
        (elem.y - expected_y).abs() < 0.1,
        "Expected y ~{expected_y}, got {}",
        elem.y
    );
    assert!(
        (elem.width - expected_w).abs() < 0.1,
        "Expected width ~{expected_w}, got {}",
        elem.width
    );
    assert!(
        (elem.height - expected_h).abs() < 0.1,
        "Expected height ~{expected_h}, got {}",
        elem.height
    );
}

#[test]
fn test_text_box_bold_formatting() {
    let runs_xml = r#"<a:r><a:rPr b="1"/><a:t>Bold text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Bold text");
    assert_eq!(para.runs[0].style.bold, Some(true));
}

#[test]
fn test_text_box_italic_formatting() {
    let runs_xml = r#"<a:r><a:rPr i="1"/><a:t>Italic text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].text, "Italic text");
    assert_eq!(para.runs[0].style.italic, Some(true));
}

#[test]
fn test_text_box_font_size() {
    let runs_xml = r#"<a:r><a:rPr sz="2400"/><a:t>Large text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs[0].style.font_size, Some(24.0));
}

/// Parse a single formatted run and return its resolved style.
fn run_style_for(runs_xml: &str) -> TextStyle {
    let shape = make_formatted_text_box(0, 0, 6_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    match &blocks[0] {
        Block::Paragraph(p) => p.runs[0].style.clone(),
        _ => panic!("Expected Paragraph"),
    }
}

#[test]
fn test_text_box_positive_character_tracking() {
    // DrawingML `spc` is hundredths of a point, like `sz`. The eyebrow run on
    // the introduction deck's title slide declares spc="220" — 2.2pt per gap.
    let runs_xml = r#"<a:r><a:rPr sz="1150" b="1" spc="220"/><a:t>OPEN SOURCE · RUST</a:t></a:r>"#;

    assert_eq!(run_style_for(runs_xml).letter_spacing, Some(2.2));
}

#[test]
fn test_text_box_negative_character_tracking() {
    // The same deck's hero wordmark tightens with spc="-100".
    let runs_xml = r#"<a:r><a:rPr sz="6800" b="1" spc="-100"/><a:t>office2pdf</a:t></a:r>"#;

    assert_eq!(run_style_for(runs_xml).letter_spacing, Some(-1.0));
}

#[test]
fn test_text_box_without_spc_has_no_tracking() {
    let runs_xml = r#"<a:r><a:rPr sz="1800"/><a:t>Untracked</a:t></a:r>"#;

    assert_eq!(run_style_for(runs_xml).letter_spacing, None);
}

#[test]
fn test_text_box_kern_threshold_is_hundredths_of_a_point() {
    // DrawingML's `kern` is the size pair kerning starts at, in hundredths of
    // a point — the unit `sz` and `spc` use. Every PowerPoint master writes
    // `kern="1200"` on its `titleStyle`, which is where the 38pt titles of
    // issue #1073 take their threshold from.
    let runs_xml = r#"<a:r><a:rPr sz="3800" kern="1200" spc="300"/><a:t>RAPPORTSTATUS</a:t></a:r>"#;

    let style = run_style_for(runs_xml);
    assert_eq!(style.pair_kerning, Some(PairKerning::AtOrAbovePt(12.0)));
    assert_eq!(style.letter_spacing, Some(3.0));
}

#[test]
fn test_text_box_kern_zero_states_never() {
    // `kern="0"` is how DrawingML records "do not kern this", not "kern from
    // 0pt up".
    let runs_xml = r#"<a:r><a:rPr sz="3800" kern="0"/><a:t>Unkerned</a:t></a:r>"#;

    assert_eq!(
        run_style_for(runs_xml).pair_kerning,
        Some(PairKerning::Never)
    );
}

#[test]
fn test_text_box_without_kern_leaves_the_decision_unstated() {
    // Absence is inheritance in DrawingML, so it must stay `None` for the
    // enclosing list style to answer.
    let runs_xml = r#"<a:r><a:rPr sz="1800"/><a:t>Inheriting</a:t></a:r>"#;

    assert_eq!(run_style_for(runs_xml).pair_kerning, None);
}

#[test]
fn test_text_box_run_baseline_preserves_positive_and_negative_offsets() {
    let runs_xml = concat!(
        r#"<a:r><a:rPr sz="1800"/><a:t>Body</a:t></a:r>"#,
        r#"<a:r><a:rPr sz="1000" baseline="30000"/><a:t>1</a:t></a:r>"#,
        r#"<a:r><a:rPr sz="1800"/><a:t> H</a:t></a:r>"#,
        r#"<a:r><a:rPr sz="1000" baseline="-25000"/><a:t>2</a:t></a:r>"#,
        r#"<a:r><a:rPr sz="1800"/><a:t>O</a:t></a:r>"#,
    );
    let shape = make_formatted_text_box(0, 0, 6_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paragraph = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph,
        other => panic!("Expected Paragraph, got {other:?}"),
    };

    assert_eq!(paragraph.runs.len(), 5);
    assert_eq!(paragraph.runs[0].style.font_size, Some(18.0));
    assert_eq!(paragraph.runs[1].style.font_size, Some(7.0));
    assert_eq!(paragraph.runs[3].style.font_size, Some(7.0));
    assert_eq!(
        paragraph.runs[1].style.baseline_shift,
        Some(BaselineShiftEm(3.0 / 7.0))
    );
    assert_eq!(
        paragraph.runs[3].style.baseline_shift,
        Some(BaselineShiftEm(-2.5 / 7.0))
    );
    assert!(paragraph.runs[0].style.baseline_shift.is_none());
    assert!(paragraph.runs[2].style.baseline_shift.is_none());
    assert!(paragraph.runs[4].style.baseline_shift.is_none());
}

#[test]
fn test_with_master_page_2_keeps_the_raised_ordinal_as_a_separate_run() {
    let data = include_bytes!("../../../../tests/fixtures/pptx/WithMaster.pptx");
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[1] {
        Page::Fixed(page) => page,
        other => panic!("Expected fixed page, got {other:?}"),
    };
    let paragraph = page
        .elements
        .iter()
        .filter_map(|element| match &element.kind {
            FixedElementKind::TextBox(text_box) => Some(text_box.content.as_slice()),
            _ => None,
        })
        .flatten()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .find(|paragraph| {
            paragraph
                .runs
                .iter()
                .map(|run| run.text.as_str())
                .collect::<String>()
                == "2nd page subtitle"
        })
        .expect("Expected the page 2 subtitle");

    assert_eq!(paragraph.runs.len(), 3);
    assert_eq!(paragraph.runs[0].text, "2");
    assert_eq!(paragraph.runs[1].text, "nd");
    assert_eq!(paragraph.runs[1].style.font_size, Some(21.0));
    assert_eq!(
        paragraph.runs[1].style.baseline_shift,
        Some(BaselineShiftEm(0.3 * 32.0 / 21.0))
    );
    assert_eq!(paragraph.runs[2].text, " page subtitle");
}

#[test]
fn test_text_box_combined_formatting() {
    let runs_xml = r#"<a:r><a:rPr b="1" i="1" u="sng" strike="sngStrike" sz="1800"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Styled text</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    let run = &para.runs[0];
    assert_eq!(run.text, "Styled text");
    assert_eq!(run.style.bold, Some(true));
    assert_eq!(run.style.italic, Some(true));
    assert_eq!(run.style.underline, Some(true));
    assert_eq!(run.style.strikethrough, Some(true));
    assert_eq!(run.style.font_size, Some(18.0));
    assert_eq!(run.style.color, Some(Color::new(255, 0, 0)));
    assert_eq!(run.style.font_family, Some("Arial".to_string()));
    assert_eq!(
        run.style.color_alpha, None,
        "a run whose fill declares no a:alpha is fully opaque"
    );
}

/// A run's `a:solidFill` colour may carry `<a:alpha>`, which PowerPoint
/// composites the ink at. The sensitivity footer a slide master stamps is the
/// common case: black at `val="50000"` reads as a half-tone, not solid black
/// (issue #1121).
#[test]
fn test_text_box_run_fill_alpha() {
    let runs_xml = r#"<a:r><a:rPr sz="800"><a:solidFill><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:rPr><a:t>Sensitivity: Internal</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    let run = &para.runs[0];
    assert_eq!(run.text, "Sensitivity: Internal");
    assert_eq!(run.style.color, Some(Color::new(0, 0, 0)));
    assert_eq!(run.style.color_alpha, Some(0.5));
}

/// Triangulation for [`test_text_box_run_fill_alpha`]: a different alpha on a
/// different colour, declared through a scheme colour rather than `srgbClr`.
#[test]
fn test_text_box_run_fill_alpha_scheme_color() {
    let runs_xml = r#"<a:r><a:rPr sz="1200"><a:solidFill><a:schemeClr val="accent1"><a:alpha val="30000"/></a:schemeClr></a:solidFill></a:rPr><a:t>Faded accent</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    let run = &para.runs[0];
    assert_eq!(run.text, "Faded accent");
    assert_eq!(run.style.color_alpha, Some(0.3));
}

#[test]
fn test_multiple_text_boxes() {
    let shape1 = make_text_box(100_000, 100_000, 2_000_000, 500_000, "Box 1");
    let shape2 = make_text_box(100_000, 700_000, 2_000_000, 500_000, "Box 2");
    let slide = make_slide_xml(&[shape1, shape2]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    assert_eq!(page.elements.len(), 2, "Expected 2 text boxes");

    let get_text = |elem: &FixedElement| -> String {
        let blocks = text_box_blocks(elem);
        blocks
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => {
                    Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                }
                _ => None,
            })
            .collect()
    };
    assert_eq!(get_text(&page.elements[0]), "Box 1");
    assert_eq!(get_text(&page.elements[1]), "Box 2");
}

#[test]
fn test_multiple_slides() {
    let slide1 = make_slide_xml(&[make_text_box(0, 0, 1_000_000, 500_000, "Slide 1")]);
    let slide2 = make_slide_xml(&[make_text_box(0, 0, 1_000_000, 500_000, "Slide 2")]);
    let slide3 = make_slide_xml(&[make_text_box(0, 0, 1_000_000, 500_000, "Slide 3")]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide1, slide2, slide3]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 3, "Expected 3 pages");
    for page in &doc.pages {
        assert!(matches!(page, Page::Fixed(_)));
    }
}

#[test]
fn test_text_box_multiple_paragraphs() {
    let paras_xml = r#"<a:p><a:r><a:rPr/><a:t>Paragraph 1</a:t></a:r></a:p><a:p><a:r><a:rPr/><a:t>Paragraph 2</a:t></a:r></a:p>"#;
    let shape = make_multi_para_text_box(0, 0, 3_000_000, 2_000_000, paras_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let paras: Vec<&Paragraph> = blocks
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => Some(p),
            _ => None,
        })
        .collect();
    assert!(paras.len() >= 2, "Expected at least 2 paragraphs");
    assert_eq!(paras[0].runs[0].text, "Paragraph 1");
    assert_eq!(paras[1].runs[0].text, "Paragraph 2");
}

#[test]
fn test_text_box_multiple_runs() {
    let runs_xml =
        r#"<a:r><a:rPr b="1"/><a:t>Bold </a:t></a:r><a:r><a:rPr i="1"/><a:t>Italic</a:t></a:r>"#;
    let shape = make_formatted_text_box(0, 0, 2_000_000, 500_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.runs.len(), 2);
    assert_eq!(para.runs[0].text, "Bold ");
    assert_eq!(para.runs[0].style.bold, Some(true));
    assert_eq!(para.runs[1].text, "Italic");
    assert_eq!(para.runs[1].style.italic, Some(true));
}

#[test]
fn test_paragraph_alignment_center() {
    let paras_xml = r#"<a:p><a:pPr algn="ctr"/><a:r><a:rPr/><a:t>Centered</a:t></a:r></a:p>"#;
    let shape = make_multi_para_text_box(0, 0, 2_000_000, 500_000, paras_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.style.alignment, Some(Alignment::Center));
}

#[test]
fn test_body_pr_vert_sets_text_rotation() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="V"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="vert"/><a:p><a:r><a:rPr lang="en-US"/><a:t>Vertical it should be!</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let slide = make_slide_xml(&[shape.to_string()]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let text_box = text_box_data(&page.elements[0]);
    assert_eq!(text_box.text_rotation_deg, Some(270.0));
}

#[test]
fn test_body_pr_vert270_sets_reverse_rotation() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="V"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="vert270"/><a:p><a:r><a:rPr lang="en-US"/><a:t>Up</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let slide = make_slide_xml(&[shape.to_string()]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let text_box = text_box_data(&page.elements[0]);
    assert_eq!(text_box.text_rotation_deg, Some(90.0));
}

#[test]
fn test_vert_text_in_rotated_pentagon_uses_text_rect_and_keeps_anchor() {
    let expected_alignments = [
        ("t", TextBoxVerticalAlign::Top),
        ("ctr", TextBoxVerticalAlign::Center),
        ("b", TextBoxVerticalAlign::Bottom),
    ];

    for (anchor, expected_alignment) in expected_alignments {
        let shape = format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="P"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="10800000"><a:off x="0" y="0"/><a:ext cx="1417740" cy="1317072"/></a:xfrm><a:prstGeom prst="pentagon"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="vert" anchor="{anchor}"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US"/><a:t>text</a:t></a:r></a:p></p:txBody></p:sp>"#
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = first_fixed_page(&doc);
        let overlay = page
            .elements
            .iter()
            .find_map(|elem| match &elem.kind {
                FixedElementKind::TextBox(tb) => Some((elem, tb)),
                _ => None,
            })
            .expect("text overlay present");
        let (element, text_box) = overlay;

        assert_eq!(text_box.text_rotation_deg, Some(270.0));
        assert_eq!(text_box.shape_rotation_deg, None);
        assert_eq!(text_box.vertical_align, expected_alignment);
        assert!((element.x - 21.320_092_374_282_623).abs() < 1e-9);
        assert!((element.y - 0.000_263_493_677_863_380_07).abs() < 1e-9);
        assert!((element.width - 68.992_886_117_576_49).abs() < 1e-9);
        assert!((element.height - 79.224_481_916_489_57).abs() < 1e-9);
        assert_eq!(text_box.padding, default_pptx_text_box_padding());
    }
}

#[test]
fn test_vert_text_in_rotated_preset_keeps_body_orientation() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="P"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="10800000"><a:off x="0" y="0"/><a:ext cx="1417740" cy="1317072"/></a:xfrm><a:prstGeom prst="pentagon"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="vert" anchor="t"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US"/><a:t>text</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let slide = make_slide_xml(&[shape.to_string()]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let background = page
        .elements
        .iter()
        .find_map(|elem| match &elem.kind {
            FixedElementKind::Shape(shape) => Some(shape),
            _ => None,
        })
        .expect("shape background present");
    let text_box = page
        .elements
        .iter()
        .find_map(|elem| match &elem.kind {
            FixedElementKind::TextBox(text_box) => Some(text_box),
            _ => None,
        })
        .expect("text overlay present");

    assert_eq!(background.rotation_deg, Some(180.0));
    assert_eq!(text_box.text_rotation_deg, Some(270.0));
    assert_eq!(
        text_box.shape_rotation_deg, None,
        "a preset background keeps the shape rotation, but its explicit vertical body keeps its own reading direction"
    );
}

#[test]
fn test_vert_text_in_angled_preset_keeps_full_box_fallback() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="P"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="2700000"><a:off x="0" y="0"/><a:ext cx="1417740" cy="1317072"/></a:xfrm><a:prstGeom prst="pentagon"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="vert" anchor="t"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US"/><a:t>text</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let slide = make_slide_xml(&[shape.to_string()]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let (element, text_box) = page
        .elements
        .iter()
        .find_map(|elem| match &elem.kind {
            FixedElementKind::TextBox(tb) => Some((elem, tb)),
            _ => None,
        })
        .expect("text overlay present");

    assert_eq!(text_box.vertical_align, TextBoxVerticalAlign::Center);
    assert_eq!(element.x, 0.0);
    assert_eq!(element.y, 0.0);
    assert!((element.width - 111.633_070_866_141_73).abs() < 1e-9);
    assert!((element.height - 103.706_456_692_913_39).abs() < 1e-9);
}

#[test]
fn test_wedge_round_rect_callout_uses_adjusted_text_rectangle() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Rounded Rectangular Callout 18"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1447800"/></a:xfrm><a:prstGeom prst="wedgeRoundRectCallout"><a:avLst><a:gd name="adj1" fmla="val 41242"/><a:gd name="adj2" fmla="val 92245"/><a:gd name="adj3" fmla="val 16667"/></a:avLst></a:prstGeom></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" b="1"/><a:t>What students should know and be able to do at each grade level and band.</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let slide = make_slide_xml(&[shape.to_string()]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let overlay = page
        .elements
        .iter()
        .find(|element| matches!(element.kind, FixedElementKind::TextBox(_)))
        .expect("text overlay present");

    // DrawingML defines the text rectangle inside each rounded corner. With
    // adj3=16667, the 19 pt corner radius contributes a 5.565 pt inset on
    // every edge; bodyPr padding is applied inside this overlay afterwards.
    let radius = 114.0 * 16_667.0 / 100_000.0;
    let inset = radius * 29_289.0 / 100_000.0;
    assert!((overlay.x - inset).abs() < 1e-9, "got {}", overlay.x);
    assert!((overlay.y - inset).abs() < 1e-9, "got {}", overlay.y);
    assert!((overlay.width - (150.0 - 2.0 * inset)).abs() < 1e-9);
    assert!((overlay.height - (114.0 - 2.0 * inset)).abs() < 1e-9);
}

#[test]
fn test_custom_geometry_uses_its_explicit_text_rectangle() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Freeform"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="1016000"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst><a:gd name="il" fmla="*/ w 10000 100000"/><a:gd name="it" fmla="*/ h 10000 100000"/><a:gd name="ir" fmla="+- r 0 il"/><a:gd name="ib" fmla="+- b 0 it"/></a:gdLst><a:rect l="il" t="it" r="ir" b="ib"/><a:pathLst><a:path><a:moveTo><a:pt x="l" y="t"/></a:moveTo><a:lnTo><a:pt x="r" y="t"/></a:lnTo><a:lnTo><a:pt x="r" y="b"/></a:lnTo><a:lnTo><a:pt x="l" y="b"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>inside</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let slide = make_slide_xml(&[shape.to_string()]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let overlay = page
        .elements
        .iter()
        .find(|element| matches!(element.kind, FixedElementKind::TextBox(_)))
        .expect("text overlay present");

    assert!((overlay.x - 10.0).abs() < 1e-9, "got {}", overlay.x);
    assert!((overlay.y - 8.0).abs() < 1e-9, "got {}", overlay.y);
    assert!((overlay.width - 80.0).abs() < 1e-9);
    assert!((overlay.height - 64.0).abs() < 1e-9);
}

#[test]
fn test_vert_text_in_plain_rect_keeps_anchor() {
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="V"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="vert" anchor="t"/><a:p><a:r><a:rPr lang="en-US"/><a:t>Up</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let slide = make_slide_xml(&[shape.to_string()]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);

    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let text_box = text_box_data(&page.elements[0]);
    assert_eq!(
        text_box.vertical_align,
        TextBoxVerticalAlign::Top,
        "plain rectangles keep their bodyPr anchor"
    );
}

#[test]
fn test_text_box_paragraph_before_after_spacing() {
    // a:spcBef / a:spcAft carry PowerPoint's inter-paragraph gaps in
    // hundredths of a point; dropping them packed bullet lists into a dense
    // block (issue #359).
    let shape = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:pPr><a:spcBef><a:spcPts val="400"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr><a:r><a:rPr lang="en-US"/><a:t>spaced bullet</a:t></a:r></a:p></p:txBody></p:sp>"#.to_string();
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let para = match &blocks[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected Paragraph"),
    };
    assert_eq!(para.style.space_before, Some(4.0));
    assert_eq!(para.style.space_after, Some(6.0));
}

// ── Hangul kinsoku break markers (issue #438) ────────────────────────
//
// PowerPoint keeps a Hangul syllable on the line when the following
// terminal punctuation overflows: it hangs the mark past the margin
// (Windows) or breaks before it (macOS). UAX #14 LB13 instead glues the
// mark to the syllable, wrapping both. The parser marks the position
// with an IR-only sentinel that the renderer turns into an empty
// `#box[]` (a UAX #14 Contingent Break), restoring the opportunity
// without touching the PDF text layer.

fn parse_single_paragraph(runs_xml: &str) -> Paragraph {
    let shape = make_formatted_text_box(0, 0, 4_000_000, 1_000_000, runs_xml);
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    match &blocks[0] {
        Block::Paragraph(p) => p.clone(),
        _ => panic!("Expected Paragraph"),
    }
}

#[test]
fn test_hangul_terminal_punct_gets_break_marker() {
    let para = parse_single_paragraph(r#"<a:r><a:rPr lang="ko-KR"/><a:t>라는 뜻인가?</a:t></a:r>"#);
    assert_eq!(para.runs[0].text, "라는 뜻인가\u{200B}?");
}

#[test]
fn test_hangul_terminal_punct_marker_spans_run_boundary() {
    // The syllable and the mark can arrive in differently-styled runs.
    let para = parse_single_paragraph(
        r#"<a:r><a:rPr lang="ko-KR" b="1"/><a:t>뜻인가</a:t></a:r><a:r><a:rPr lang="ko-KR"/><a:t>?</a:t></a:r>"#,
    );
    assert_eq!(para.runs[0].text, "뜻인가");
    assert_eq!(para.runs[1].text, "\u{200B}?");
}

#[test]
fn test_hangul_break_marker_covers_fullwidth_and_ellipsis() {
    let para = parse_single_paragraph(
        r#"<a:r><a:rPr lang="ko-KR"/><a:t>질문？최고！끝。쉼、다음…</a:t></a:r>"#,
    );
    assert_eq!(
        para.runs[0].text,
        "질문\u{200B}？최고\u{200B}！끝\u{200B}。쉼\u{200B}、다음\u{200B}…"
    );
}

#[test]
fn test_no_break_marker_for_percent_or_latin_context() {
    // '%' glues to the syllable in PowerPoint (measured in #438), and a
    // Latin letter before the mark is outside the rule entirely.
    let para = parse_single_paragraph(
        r#"<a:r><a:rPr lang="ko-KR"/><a:t>확률이 3%라는 것. OK?</a:t></a:r>"#,
    );
    assert_eq!(para.runs[0].text, "확률이 3%라는 것\u{200B}. OK?");
}

#[test]
fn test_no_break_marker_between_hangul_syllables() {
    let para = parse_single_paragraph(r#"<a:r><a:rPr lang="ko-KR"/><a:t>가나다라마</a:t></a:r>"#);
    assert_eq!(para.runs[0].text, "가나다라마");
}

// ----- Text fields, `<a:fld>` (issue #540) -----

/// A slide holding one text box whose paragraph is a single `<a:fld>`.
fn slide_with_field(field_xml: &str) -> String {
    make_slide_xml(&[format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="SlideNumber"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="500000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p>{field_xml}</a:p></p:txBody></p:sp>"#
    )])
}

/// The text of the first run on `page_index`, or `None` when the slide has none.
fn field_run_text(slides: &[String], page_index: usize) -> Option<String> {
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, slides);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let page = match &doc.pages[page_index] {
        Page::Fixed(page) => page,
        _ => panic!("expected a fixed page"),
    };
    let blocks = text_box_blocks(&page.elements[0]);
    match blocks.first()? {
        Block::Paragraph(paragraph) => Some(paragraph.runs.first()?.text.clone()),
        _ => None,
    }
}

#[test]
fn a_slidenum_field_renders_the_decks_own_position() {
    // PowerPoint caches the number it last drew inside the field. That cache
    // goes stale when slides are reordered, so the deck position wins.
    let slides = vec![
        slide_with_field(
            r#"<a:fld id="{F7021451-0000-0000-0000-000000000000}" type="slidenum"><a:rPr lang="en-US"/><a:t>7</a:t></a:fld>"#,
        ),
        slide_with_field(
            r#"<a:fld id="{F7021452-0000-0000-0000-000000000000}" type="slidenum"><a:rPr lang="en-US"/><a:t>7</a:t></a:fld>"#,
        ),
    ];

    assert_eq!(field_run_text(&slides, 0).as_deref(), Some("1"));
    assert_eq!(field_run_text(&slides, 1).as_deref(), Some("2"));
}

#[test]
fn an_unhandled_field_type_falls_back_to_its_cached_text() {
    // PowerPoint's own fallback, and the only deterministic reading of a
    // date field: the text the authoring application last wrote.
    let slides = vec![slide_with_field(
        r#"<a:fld id="{F7021453-0000-0000-0000-000000000000}" type="datetime1"><a:rPr lang="en-US"/><a:t>2026-06-01</a:t></a:fld>"#,
    )];

    assert_eq!(field_run_text(&slides, 0).as_deref(), Some("2026-06-01"));
}

#[test]
fn a_field_carries_its_own_run_properties() {
    let slides = vec![slide_with_field(
        r#"<a:fld id="{F7021454-0000-0000-0000-000000000000}" type="slidenum"><a:rPr lang="en-US" b="1" sz="1400"/><a:t>1</a:t></a:fld>"#,
    )];
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &slides);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let run = match &blocks[0] {
        Block::Paragraph(paragraph) => &paragraph.runs[0],
        _ => panic!("expected a paragraph"),
    };

    assert_eq!(run.style.bold, Some(true));
    assert_eq!(run.style.font_size, Some(14.0));
}

#[test]
fn a_field_beside_literal_runs_keeps_the_reading_order() {
    let slides = vec![slide_with_field(
        r#"<a:r><a:rPr lang="en-US"/><a:t>Page </a:t></a:r><a:fld id="{F7021455-0000-0000-0000-000000000000}" type="slidenum"><a:rPr lang="en-US"/><a:t>9</a:t></a:fld><a:r><a:rPr lang="en-US"/><a:t> of 1</a:t></a:r>"#,
    )];
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &slides);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    let text: String = match &blocks[0] {
        Block::Paragraph(paragraph) => paragraph.runs.iter().map(|run| run.text.as_str()).collect(),
        _ => panic!("expected a paragraph"),
    };

    assert_eq!(text, "Page 1 of 1");
}

/// `a:rPr/@cap` is how PowerPoint uppercases a run at render time — the text
/// stays mixed-case in the file. The deck on issue #875 uses it 25 times, and
/// its closing title reads `SPØRSMÅL OG SVAR` in every other renderer while we
/// printed the stored `Spørsmål og svar`.
fn text_box_with_cap(cap: &str, text: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="1000000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" cap="{cap}"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

fn first_run_style_of(slide_xml: &str) -> TextStyle {
    let data = build_test_pptx(
        SLIDE_CX,
        SLIDE_CY,
        &[make_slide_xml(&[slide_xml.to_string()])],
    );
    let (doc, _warnings) = PptxParser
        .parse(&data, &ConvertOptions::default())
        .expect("parses");
    let page = first_fixed_page(&doc);
    let blocks = text_box_blocks(&page.elements[0]);
    match &blocks[0] {
        Block::Paragraph(p) => p.runs[0].style.clone(),
        other => panic!("expected a paragraph, got {other:?}"),
    }
}

#[test]
fn test_run_cap_all_sets_all_caps() {
    let style = first_run_style_of(&text_box_with_cap("all", "Spørsmål og svar"));
    assert_eq!(style.all_caps, Some(true));
    assert_ne!(style.small_caps, Some(true));
}

#[test]
fn test_run_cap_small_sets_small_caps() {
    let style = first_run_style_of(&text_box_with_cap("small", "Spørsmål og svar"));
    assert_eq!(style.small_caps, Some(true));
    assert_ne!(style.all_caps, Some(true));
}

/// `cap="none"` is an explicit "do not case this run", which PowerPoint writes
/// to override an inherited `cap`. It must state the answer, not stay silent.
#[test]
fn test_run_cap_none_states_no_casing() {
    let style = first_run_style_of(&text_box_with_cap("none", "Spørsmål og svar"));
    assert_eq!(style.all_caps, Some(false));
    assert_eq!(style.small_caps, Some(false));
}

/// A run that declares no `cap` leaves both unset, so an inherited answer still
/// stands.
#[test]
fn test_a_run_without_cap_leaves_casing_unset() {
    let shape = make_text_box(0, 0, 4_000_000, 1_000_000, "Spørsmål og svar");
    let style = first_run_style_of(&shape);
    assert_eq!(style.all_caps, None);
    assert_eq!(style.small_caps, None);
}

/// `<a:spAutoFit/>` and `<a:normAutofit/>` are different requests
/// (ECMA-376 §21.1.2.1.2 / §21.1.2.1.3) and only the second one shrinks text
/// (issue #898).
///
/// - *shape autofit* grows the **shape** to the text; the run keeps its
///   declared size, and PowerPoint saves the grown box.
/// - *normal autofit* shrinks the **text** on overflow, by the `fontScale`
///   and `lnSpcReduction` it states.
///
/// The deck in #841 puts `<a:spAutoFit/>` on a 9.6pt-tall box holding 8pt
/// text, and we scaled the run to 4.9pt to make it fit one line.
#[test]
fn sp_auto_fit_does_not_shrink_text_but_norm_autofit_does() {
    for (autofit, expected_scaling) in [("<a:spAutoFit/>", false), ("<a:normAutofit/>", true)] {
        let shape = format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="63500" y="6672580"/><a:ext cx="855663" cy="121920"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr horzOverflow="overflow" lIns="0" tIns="0" rIns="0" bIns="0">{autofit}</a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="800"/><a:t>Sensitivity: Internal</a:t></a:r></a:p></p:txBody></p:sp>"#
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = first_fixed_page(&doc);
        let FixedElementKind::TextBox(ref text_box) = page.elements[0].kind else {
            panic!("expected a text box");
        };
        assert_eq!(
            text_box.auto_fit,
            expected_scaling,
            "{autofit} must {} request text scaling",
            if expected_scaling { "" } else { "not" }
        );
    }
}

fn parse_normal_autofit_text_box(attributes: &str, paragraphs_xml: &str) -> TextBoxData {
    let shape = format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr><a:normAutofit {attributes}/></a:bodyPr><a:lstStyle/>{paragraphs_xml}</p:txBody></p:sp>"#
    );
    let slide = make_slide_xml(&[shape]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let Page::Fixed(mut page) = doc.pages.into_iter().next().expect("one slide") else {
        panic!("expected a fixed page");
    };
    let FixedElementKind::TextBox(text_box) = page.elements.remove(0).kind else {
        panic!("expected a text box");
    };
    text_box
}

/// PowerPoint paints a saved `fontScale` at the nearest whole point, half up:
/// a native export of the #1303 fixture sets its 36pt title at 32pt, not
/// 36 x 0.9 = 32.4pt, and the #1375 deck's 22pt body at 15pt, not 15.4pt.
#[test]
fn normal_autofit_applies_powerpoints_saved_font_scale() {
    for (attributes, expected_font_size) in [
        (r#"fontScale="90000""#, 32.0),
        (r#"fontScale="75.000%""#, 27.0),
    ] {
        let text_box = parse_normal_autofit_text_box(
            attributes,
            r#"<a:p><a:r><a:rPr sz="3600"/><a:t>Scaled title</a:t></a:r></a:p>"#,
        );
        let Block::Paragraph(paragraph) = &text_box.content[0] else {
            panic!("expected a paragraph");
        };

        assert_eq!(paragraph.runs[0].style.font_size, Some(expected_font_size));
        assert!(
            !text_box.auto_fit,
            "a saved fontScale is PowerPoint's completed answer, not a request to shrink again"
        );
    }
}

/// One-factor native probe on the #1375 deck (`fontScale` varied on the
/// slide-12 body placeholder, 22pt body and 35pt heading, macOS PowerPoint
/// 16.x, Quartz PDFContext): every painted size is the scaled size rounded
/// half up to a whole point, including the exact halves 24.5 -> 25 and
/// 16.5 -> 17 that round-half-even or truncation would send the other way.
#[test]
fn normal_autofit_rounds_the_scaled_size_half_up_to_a_whole_point() {
    for (font_scale, body_sz, heading_sz, expected_body, expected_heading) in [
        ("62500", "2200", "3500", 14.0, 22.0),
        ("70000", "2200", "3500", 15.0, 25.0),
        ("75000", "2200", "3500", 17.0, 26.0),
        ("80000", "2200", "3500", 18.0, 28.0),
        ("85000", "2200", "3500", 19.0, 30.0),
        ("92500", "2200", "3500", 20.0, 32.0),
        ("97500", "2200", "3500", 21.0, 34.0),
        ("90000", "3600", "1050", 32.0, 9.0),
    ] {
        let text_box = parse_normal_autofit_text_box(
            &format!(r#"fontScale="{font_scale}""#),
            &format!(
                concat!(
                    r#"<a:p><a:r><a:rPr sz="{heading_sz}" b="1"/><a:t>Heading</a:t></a:r></a:p>"#,
                    r#"<a:p><a:r><a:rPr sz="{body_sz}" i="1"/><a:t>Body</a:t></a:r></a:p>"#,
                ),
                heading_sz = heading_sz,
                body_sz = body_sz,
            ),
        );
        let sizes: Vec<Option<f64>> = text_box
            .content
            .iter()
            .map(|block| match block {
                Block::Paragraph(paragraph) => paragraph.runs[0].style.font_size,
                other => panic!("expected a paragraph, got {other:?}"),
            })
            .collect();

        assert_eq!(
            sizes,
            vec![Some(expected_heading), Some(expected_body)],
            "fontScale={font_scale} on sz={heading_sz}/{body_sz}"
        );
    }
}

#[test]
fn normal_autofit_scales_explicit_list_marker_sizes_too() {
    let text_box = parse_normal_autofit_text_box(
        r#"fontScale="90000""#,
        r#"<a:p><a:pPr><a:buSzPts val="1500"/><a:buChar char="•"/></a:pPr><a:r><a:rPr sz="2000"/><a:t>Scaled bullet</a:t></a:r></a:p>"#,
    );
    let Block::List(list) = &text_box.content[0] else {
        panic!("expected a list");
    };

    assert_eq!(list.items[0].content[0].runs[0].style.font_size, Some(18.0));
    assert_eq!(
        list.level_styles[&0]
            .marker_style
            .as_ref()
            .and_then(|style| style.font_size),
        Some(14.0),
        "a 15pt bullet at 90% is 13.5pt, which PowerPoint paints at 14pt"
    );
}

/// ECMA-376 21.1.2.1.2 subtracts `lnSpcReduction` from the percentage line
/// spacing, so 120% less 20% is 100%. The multiplicative reading (120% x 0.8 =
/// 96%) paced the #1375 deck's 110%-less-20% body at 88% where PowerPoint
/// paints 90%, the same pitch as its 90% neighbours.
#[test]
fn normal_autofit_reduces_only_percentage_line_spacing() {
    let text_box = parse_normal_autofit_text_box(
        r#"fontScale="80000" lnSpcReduction="20.000%""#,
        concat!(
            r#"<a:p><a:pPr><a:lnSpc><a:spcPct val="120000"/></a:lnSpc></a:pPr><a:r><a:rPr sz="2000"/><a:t>Percentage</a:t></a:r></a:p>"#,
            r#"<a:p><a:pPr><a:lnSpc><a:spcPts val="1800"/></a:lnSpc></a:pPr><a:r><a:rPr sz="2000"/><a:t>Exact</a:t></a:r></a:p>"#,
            r#"<a:p><a:r><a:rPr sz="2000"/><a:t>Default percentage</a:t></a:r></a:p>"#,
        ),
    );

    let paragraphs: Vec<&Paragraph> = text_box
        .content
        .iter()
        .map(|block| match block {
            Block::Paragraph(paragraph) => paragraph,
            other => panic!("expected a paragraph, got {other:?}"),
        })
        .collect();
    assert!(matches!(
        paragraphs[0].style.line_spacing,
        Some(LineSpacing::Proportional(value)) if (value - 1.0).abs() < 1e-9
    ));
    assert!(matches!(
        paragraphs[1].style.line_spacing,
        Some(LineSpacing::Exact(value)) if (value - 18.0).abs() < 1e-9
    ));
    assert!(matches!(
        paragraphs[2].style.line_spacing,
        Some(LineSpacing::Proportional(value)) if (value - 0.8).abs() < 1e-9
    ));
    for paragraph in &paragraphs {
        assert_eq!(paragraph.runs[0].style.font_size, Some(16.0));
    }
    assert!(!text_box.auto_fit);
}

/// `lnSpcReduction` reduces percentage line spacing inside each paragraph,
/// not a separate percentage gap before it. The gap stays 20% of the plain
/// 1.2em line: 0.2 * 1.2 * 20pt = 4.8pt (#1300, #1343).
#[test]
fn normal_autofit_does_not_reduce_percentage_paragraph_spacing() {
    let text_box = parse_normal_autofit_text_box(
        r#"lnSpcReduction="10.000%""#,
        concat!(
            r#"<a:p><a:pPr><a:spcAft><a:spcPct val="20000"/></a:spcAft></a:pPr><a:r><a:rPr sz="2000"/><a:t>First</a:t></a:r></a:p>"#,
            r#"<a:p><a:pPr><a:spcBef><a:spcPct val="20000"/></a:spcBef></a:pPr><a:r><a:rPr sz="2000"/><a:t>Second</a:t></a:r></a:p>"#,
        ),
    );
    let Block::Paragraph(first) = &text_box.content[0] else {
        panic!("expected the first paragraph")
    };
    let Block::Paragraph(second) = &text_box.content[1] else {
        panic!("expected the second paragraph")
    };

    assert!(matches!(
        second.style.line_spacing,
        Some(LineSpacing::Proportional(value)) if (value - 0.9).abs() < 1e-9
    ));
    assert!((first.style.space_after.unwrap() - 4.8).abs() < 1e-9);
    assert!((second.style.space_before.unwrap() - 4.8).abs() < 1e-9);
    assert_eq!(first.style.space_after_percent, None);
    assert_eq!(second.style.space_before_percent, None);
}
