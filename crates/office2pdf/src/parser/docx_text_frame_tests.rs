//! A body paragraph in a `w:framePr` frame hung `w:y` below its anchor
//! paragraph: Word lifts it out of the flow, draws it in its frame, and lays
//! the text after it beside the frame. Empty paragraphs in a frame at the top
//! margin keep the frame's box there, which a full-width table clears. A
//! frame at a fixed place on its page that keeps text off its sides is drawn
//! there, and the body text that meets it starts below it (issue #2231).

use super::*;

const MARGIN_PT: f64 = 72.0;
/// A4's 11906 twips less two 1440-twip margins.
const TEXT_WIDTH_PT: f64 = 451.3;
/// The frame width every frame here states: 3000 twips.
const FRAME_WIDTH_PT: f64 = 150.0;
/// The `w:hSpace` the wrapping tests state: 180 twips.
const FRAME_SPACE_PT: f64 = 9.0;

fn run(content: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>{content}</w:r>"#
    )
}

fn paragraph(paragraph_properties: &str, text: &str) -> String {
    format!(
        r#"<w:p><w:pPr>{paragraph_properties}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>{}</w:p>"#,
        run(&format!("<w:t>{text}</w:t>"))
    )
}

/// A 3000-twip-wide wrapping frame `height_twips` tall, hung `y_twips` below
/// its anchor paragraph; `extra` states the rest of its attributes.
fn offset_frame(height_twips: u32, y_twips: i32, extra: &str) -> String {
    format!(
        r#"<w:framePr w:w="3000" w:h="{height_twips}" w:wrap="around" w:vAnchor="text" w:y="{y_twips}" {extra}/>"#
    )
}

/// An A4 section's page size and 72pt margins.
const PAGE_SETUP: &str = r#"<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/>"#;

/// An A4 package with 72pt margins, no settings part and no styles part.
fn frame_docx(body: &str) -> Vec<u8> {
    section_docx(body, PAGE_SETUP)
}

/// [`frame_docx`] with `section_properties` in the body's closing
/// `w:sectPr`.
fn section_docx(body: &str, section_properties: &str) -> Vec<u8> {
    package_docx(body, section_properties, None)
}

/// [`frame_docx`] whose settings declare `compatibilityMode` `mode`.
fn compat_frame_docx(body: &str, mode: u32) -> Vec<u8> {
    package_docx(body, PAGE_SETUP, Some(mode))
}

fn package_docx(body: &str, section_properties: &str, compatibility_mode: Option<u32>) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr>{section_properties}</w:sectPr></w:body></w:document>"#
    );
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    if let Some(mode) = compatibility_mode {
        zip.start_file("word/settings.xml", options).unwrap();
        let settings: String = format!(
            r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{mode}"/></w:compat></w:settings>"#
        );
        std::io::Write::write_all(&mut zip, settings.as_bytes()).unwrap();
    }
    let parts: [(&str, &str); 3] = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        ("word/document.xml", &document),
    ];
    for (name, content) in parts {
        zip.start_file(name, options).unwrap();
        std::io::Write::write_all(&mut zip, content.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// The painted non-blank runs, one per line, with each run's advance width.
#[cfg(not(target_arch = "wasm32"))]
fn placed(data: &[u8]) -> Vec<(crate::render::pdf::PlacedTextRun, f64)> {
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    let text = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
    let glyphs = crate::render::pdf::compiled_glyph_runs(&source.source, 0, true).unwrap();
    text.into_iter()
        .filter(|run| !run.text.trim().is_empty())
        .map(|run| {
            let width = glyphs
                .iter()
                .find(|glyph| glyph.text == run.text)
                .unwrap_or_else(|| panic!("no glyph run {:?}", run.text))
                .width_pt();
            (run, width)
        })
        .collect()
}

#[cfg(not(target_arch = "wasm32"))]
fn find<'a>(
    runs: &'a [(crate::render::pdf::PlacedTextRun, f64)],
    text: &str,
) -> &'a (crate::render::pdf::PlacedTextRun, f64) {
    runs.iter()
        .find(|(run, _)| run.text.trim_end() == text)
        .unwrap_or_else(|| panic!("no run {text:?} in {runs:?}"))
}

#[cfg(not(target_arch = "wasm32"))]
fn assert_close(actual: f64, expected: f64, context: &str) {
    assert!(
        (actual - expected).abs() < 0.05,
        "{context}: {actual} != {expected}"
    );
}

/// The body blocks of the package's only section.
fn body_blocks(data: &[u8]) -> Vec<Block> {
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let Page::Flow(page) = &document.pages[0] else {
        panic!("expected a flow page");
    };
    page.content.clone()
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn offset_frame_is_drawn_below_its_anchor_and_leaves_no_line() {
    let control = placed(&frame_docx(
        &[
            paragraph("", "Anchor line"),
            paragraph("", "Following text"),
        ]
        .concat(),
    ));
    let (control_following, _) = find(&control, "Following text");
    for (h_anchor, x_twips, origin_pt) in [
        ("page", 2000_i32, 0.0),
        ("margin", 600, MARGIN_PT),
        ("text", 1200, MARGIN_PT),
    ] {
        for y_twips in [800_i32, 1400] {
            let frame = offset_frame(
                600,
                y_twips,
                &format!(r#"w:hRule="exact" w:hAnchor="{h_anchor}" w:x="{x_twips}""#),
            );
            let data = frame_docx(
                &[
                    paragraph("", "Anchor line"),
                    paragraph(&frame, "Framed text"),
                    paragraph("", "Following text"),
                ]
                .concat(),
            );
            let runs = placed(&data);
            let context = format!("hAnchor={h_anchor} x={x_twips} y={y_twips}");
            // The frame takes no line: the next paragraph starts where the
            // framed one would have, and its line ends above the frame.
            let (following, _) = find(&runs, "Following text");
            assert_close(
                following.baseline_pt,
                control_following.baseline_pt,
                &context,
            );
            assert_close(following.left_pt, MARGIN_PT, &context);
            // The frame hangs w:y below that same point, w:x from its anchor.
            let (framed, _) = find(&runs, "Framed text");
            assert_close(
                framed.left_pt,
                origin_pt + f64::from(x_twips) / 20.0,
                &context,
            );
            assert_close(
                framed.baseline_pt,
                following.baseline_pt + f64::from(y_twips) / 20.0,
                &context,
            );
        }
    }
}

/// A frame's content starts at its top, though the section around it is
/// centred vertically on the page.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn offset_frame_content_starts_at_its_top_in_a_vertically_centred_section() {
    let y_twips: i32 = 800;
    let centred: String = format!(r#"{PAGE_SETUP}<w:vAlign w:val="center"/>"#);
    for height_twips in [1000_u32, 2400] {
        let frame = offset_frame(
            height_twips,
            y_twips,
            r#"w:hRule="exact" w:hAnchor="margin" w:x="0""#,
        );
        let data = section_docx(
            &[
                paragraph("", "Anchor line"),
                paragraph(&frame, "Framed text"),
                paragraph("", "Following text"),
            ]
            .concat(),
            &centred,
        );
        let runs = placed(&data);
        let context = format!("h={height_twips}");
        let (anchor, _) = find(&runs, "Anchor line");
        assert!(
            anchor.baseline_pt > 300.0,
            "{context}: the section must be centred: {runs:?}"
        );
        let (following, _) = find(&runs, "Following text");
        let (framed, _) = find(&runs, "Framed text");
        assert_close(
            framed.baseline_pt,
            following.baseline_pt + f64::from(y_twips) / 20.0,
            &context,
        );
    }
}

/// The frame hangs from the end of the gap after the paragraph before it;
/// the gap before the paragraph after it, which Word lays only between the
/// two flowing paragraphs, does not move it.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn an_offset_frame_hangs_below_the_gap_after_the_paragraph_before_it() {
    let y_twips: i32 = 1000;
    let spaced = |before: u32, after: u32, text: &str| {
        format!(
            r#"<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}" w:line="240" w:lineRule="auto"/></w:pPr>{}</w:p>"#,
            run(&format!("<w:t>{text}</w:t>"))
        )
    };
    // The following paragraph can also open a continuous section of its
    // own, as in the issue's case, where the framed paragraph ends the first.
    let section_end: String = format!("<w:sectPr>{PAGE_SETUP}</w:sectPr>");
    let continuous: String = format!(r#"<w:type w:val="continuous"/>{PAGE_SETUP}"#);
    for ends_section in [false, true] {
        for (after_twips, before_twips) in [(0_u32, 240_u32), (120, 240), (240, 120)] {
            let frame = offset_frame(
                600,
                y_twips,
                r#"w:hRule="exact" w:hAnchor="margin" w:x="0""#,
            );
            let frame_properties: String = if ends_section {
                format!("{frame}{section_end}")
            } else {
                frame
            };
            let data = section_docx(
                &[
                    spaced(0, after_twips, "Anchor line"),
                    format!(
                        r#"<w:p><w:pPr>{frame_properties}</w:pPr>{}</w:p>"#,
                        run("<w:t>Framed text</w:t>")
                    ),
                    spaced(before_twips, 0, "Following text"),
                ]
                .concat(),
                if ends_section {
                    &continuous
                } else {
                    PAGE_SETUP
                },
            );
            let runs = placed(&data);
            let context =
                format!("after={after_twips} before={before_twips} ends section={ends_section}");
            let (following, _) = find(&runs, "Following text");
            let (framed, _) = find(&runs, "Framed text");
            let (after, before): (f64, f64) = (
                f64::from(after_twips) / 20.0,
                f64::from(before_twips) / 20.0,
            );
            // The following line starts the larger of the two gaps below the
            // anchor line; the frame, the first gap and w:y below it.
            assert_close(
                framed.baseline_pt,
                following.baseline_pt - after.max(before) + after + f64::from(y_twips) / 20.0,
                &context,
            );
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn text_meeting_an_offset_frame_lies_beside_it() {
    let lines: Vec<String> = (1..=8)
        .map(|line| paragraph("", &format!("Line {line}")))
        .collect();
    let right_lines: Vec<String> = (1..=8)
        .map(|line| paragraph(r#"<w:jc w:val="right"/>"#, &format!("Line {line}")))
        .collect();
    // A frame at the column's left edge leaves its room on the right, one at
    // the right edge its room on the left.
    let frame_right_x: f64 = TEXT_WIDTH_PT - FRAME_WIDTH_PT;
    for (x_twips, body) in [(0.0, &lines), (frame_right_x, &right_lines)] {
        for height_twips in [700_u32, 1100] {
            let frame = offset_frame(
                height_twips,
                20,
                &format!(
                    r#"w:hRule="exact" w:hSpace="180" w:hAnchor="margin" w:x="{}""#,
                    x_twips * 20.0
                ),
            );
            let data = frame_docx(
                &[
                    paragraph("", "Anchor line"),
                    paragraph(&frame, "Framed text"),
                    body.concat(),
                ]
                .concat(),
            );
            let runs = placed(&data);
            let context = format!("frame x={x_twips} h={height_twips}");
            let (first, _) = find(&runs, "Line 1");
            // The frame hangs 1pt below the first line's top.
            let band_bottom: f64 =
                first.baseline_pt - first.font_size_pt + 1.0 + f64::from(height_twips) / 20.0;
            let (mut beside, mut below) = (0, 0);
            for (run, width) in runs.iter().filter(|(run, _)| run.text.starts_with("Line")) {
                let (left, right): (f64, f64) = (run.left_pt, run.left_pt + width);
                if run.baseline_pt < band_bottom {
                    if x_twips == 0.0 {
                        assert_close(left, MARGIN_PT + FRAME_WIDTH_PT + FRAME_SPACE_PT, &context);
                    } else {
                        assert_close(right, MARGIN_PT + frame_right_x - FRAME_SPACE_PT, &context);
                    }
                    beside += 1;
                } else if run.baseline_pt - run.font_size_pt > band_bottom {
                    if x_twips == 0.0 {
                        assert_close(left, MARGIN_PT, &context);
                    } else {
                        assert_close(right, MARGIN_PT + TEXT_WIDTH_PT, &context);
                    }
                    below += 1;
                }
            }
            assert!(
                beside > 0 && below > 0,
                "{context}: {beside} lines beside, {below} below in {runs:?}"
            );
        }
    }
}

/// An exact frame keeps text off the height it states; an at-least one off
/// its content when that is taller.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn an_offset_frame_excludes_its_stated_or_its_content_height() {
    let framed_text: &str =
        "Framed words that need several lines in a narrow frame, which is what this paragraph does";
    let lines: Vec<String> = (1..=8)
        .map(|line| paragraph("", &format!("Line {line}")))
        .collect();
    for rule in ["exact", "atLeast"] {
        let frame = offset_frame(
            200,
            20,
            &format!(r#"w:hRule="{rule}" w:hAnchor="margin" w:x="0""#),
        );
        let data = frame_docx(
            &[
                paragraph("", "Anchor line"),
                paragraph(&frame, framed_text),
                lines.concat(),
            ]
            .concat(),
        );
        let runs = placed(&data);
        let (first, _) = find(&runs, "Line 1");
        let frame_top: f64 = first.baseline_pt - first.font_size_pt + 1.0;
        let content_bottom: f64 = runs
            .iter()
            .filter(|(run, _)| framed_text.contains(run.text.trim()))
            .map(|(run, _)| run.baseline_pt)
            .fold(f64::MIN, f64::max);
        assert!(
            content_bottom > frame_top + 30.0,
            "{rule}: the framed text must take several lines: {runs:?}"
        );
        // Line tops, like the frame top, are taken an em above the baseline.
        // An exact frame ends 10pt (200 twips) down; an at-least one below
        // its last line, whose box reaches some way past that baseline.
        let (beside_before, below_after): (f64, f64) = match rule {
            "exact" => (frame_top + 9.5, frame_top + 10.5),
            _ => (content_bottom, content_bottom + 0.5 * first.font_size_pt),
        };
        let (mut beside, mut below) = (0, 0);
        for (run, _) in runs.iter().filter(|(run, _)| run.text.starts_with("Line")) {
            let line_top: f64 = run.baseline_pt - run.font_size_pt;
            if line_top < beside_before {
                assert_close(run.left_pt, MARGIN_PT + FRAME_WIDTH_PT, rule);
                beside += 1;
            } else if line_top > below_after {
                assert_close(run.left_pt, MARGIN_PT, rule);
                below += 1;
            }
        }
        assert!(
            beside > 0 && below > 0,
            "{rule}: {beside} lines beside, {below} below in {runs:?}"
        );
    }
}

/// A paragraph whose first line ends above the frame but which reaches it
/// starts below the frame rather than run over it; one that ends above the
/// frame stays where it is.
///
/// TODO(per-line wrap beside a float, #2355): Word keeps the first lines
/// above the frame and narrows the ones beside it.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_paragraph_reaching_an_offset_frame_from_above_starts_below_it() {
    // Ten lines, which reach both frames; the short paragraph is one line.
    let long_text: String = "Long paragraph words ".repeat(40);
    for y_twips in [1000_i32, 1400] {
        for (text, reaches_frame) in [("Short paragraph", false), (long_text.trim_end(), true)] {
            let first_words: String = text.split(' ').take(2).collect::<Vec<_>>().join(" ");
            let first_line = |runs: &[(crate::render::pdf::PlacedTextRun, f64)]| {
                runs.iter()
                    .find(|(run, _)| run.text.starts_with(&first_words))
                    .map(|(run, _)| run.clone())
                    .unwrap_or_else(|| panic!("no line starting {first_words:?} in {runs:?}"))
            };
            let control = placed(&frame_docx(
                &[paragraph("", "Anchor line"), paragraph("", text)].concat(),
            ));
            let frame = offset_frame(
                600,
                y_twips,
                r#"w:hRule="exact" w:hAnchor="margin" w:x="0""#,
            );
            let data = frame_docx(
                &[
                    paragraph("", "Anchor line"),
                    paragraph(&frame, "Framed text"),
                    paragraph("", text),
                ]
                .concat(),
            );
            let runs = placed(&data);
            let context = format!("y={y_twips} reaches frame={reaches_frame}");
            let line = first_line(&runs);
            let expected: f64 = first_line(&control).baseline_pt
                + if reaches_frame {
                    f64::from(y_twips) / 20.0 + 30.0
                } else {
                    0.0
                };
            assert_close(line.baseline_pt, expected, &context);
            assert_close(line.left_pt, MARGIN_PT, &context);
        }
    }
}

/// A frame that ends its section still keeps the text of a continuous
/// section after it off its extent, while that section is laid out in the
/// same text area.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn an_offset_frame_excludes_text_in_the_continuous_section_after_it() {
    let lines: Vec<String> = (1..=8)
        .map(|line| paragraph("", &format!("Line {line}")))
        .collect();
    let frame = offset_frame(700, 20, r#"w:hRule="exact" w:hAnchor="margin" w:x="0""#);
    let continuous: &str = r#"<w:type w:val="continuous"/>"#;
    // The second section keeps the first's margins, or moves its left one in
    // to 144pt (2880 twips), which leaves the frame's band in another area.
    for (second_left_twips, wraps) in [(1440_u32, true), (2880, false)] {
        let body: String = [
            paragraph("", "Anchor line"),
            format!(
                r#"<w:p><w:pPr>{frame}<w:sectPr>{PAGE_SETUP}</w:sectPr></w:pPr>{}</w:p>"#,
                run("<w:t>Framed text</w:t>")
            ),
            lines.concat(),
        ]
        .concat();
        let second_section: String = format!(
            r#"{continuous}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="{second_left_twips}" w:right="1440"/>"#
        );
        let runs = placed(&section_docx(&body, &second_section));
        let context = format!("second section left={second_left_twips}");
        let margin: f64 = f64::from(second_left_twips) / 20.0;
        let (first, _) = find(&runs, "Line 1");
        if wraps {
            assert_close(first.left_pt, MARGIN_PT + FRAME_WIDTH_PT, &context);
        } else {
            assert_close(first.left_pt, margin, &context);
        }
        let (last, _) = find(&runs, "Line 8");
        assert_close(last.left_pt, margin, &context);
    }
}

#[test]
fn only_a_wrapping_text_anchored_frame_with_height_and_offset_is_lifted() {
    let lifted: String = offset_frame(600, 400, r#"w:hRule="exact" w:hAnchor="margin" w:x="0""#);
    let kept: [(&str, String); 8] = [
        (
            "no height",
            r#"<w:framePr w:w="3000" w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:x="0" w:y="400"/>"#.to_string(),
        ),
        (
            "no width",
            r#"<w:framePr w:h="600" w:hRule="exact" w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:x="0" w:y="400"/>"#.to_string(),
        ),
        ("no offset", offset_frame(600, 0, r#"w:hRule="exact""#)),
        (
            "aligned vertically",
            offset_frame(600, 400, r#"w:hRule="exact" w:yAlign="top""#),
        ),
        (
            "not beside",
            r#"<w:framePr w:w="3000" w:h="600" w:hRule="exact" w:wrap="notBeside" w:vAnchor="text" w:y="400"/>"#.to_string(),
        ),
        (
            "no wrap",
            r#"<w:framePr w:w="3000" w:h="600" w:hRule="exact" w:wrap="none" w:vAnchor="text" w:y="400"/>"#.to_string(),
        ),
        (
            "margin anchored",
            r#"<w:framePr w:w="3000" w:h="600" w:hRule="exact" w:wrap="around" w:vAnchor="margin" w:y="400"/>"#.to_string(),
        ),
        (
            "page anchored",
            r#"<w:framePr w:w="3000" w:h="600" w:hRule="exact" w:wrap="around" w:vAnchor="page" w:y="400"/>"#.to_string(),
        ),
    ];
    let is_frame = |block: &Block| matches!(block, Block::TextFrame(_));
    let blocks = body_blocks(&frame_docx(&paragraph(&lifted, "Framed text")));
    let [Block::TextFrame(frame)] = blocks.as_slice() else {
        panic!("expected one text frame: {blocks:?}");
    };
    assert!(
        matches!(frame.content.as_slice(), [Block::Paragraph(paragraph)] if paragraph.runs[0].text == "Framed text"),
        "{frame:?}"
    );
    for (case, frame) in kept {
        let blocks = body_blocks(&frame_docx(&paragraph(&frame, "Framed text")));
        assert!(!blocks.iter().any(is_frame), "{case}: {blocks:?}");
    }
    // Adjacent paragraphs with an identical frame share it; this path lifts
    // a frame of one paragraph only.
    let shared = body_blocks(&frame_docx(
        &[
            paragraph(&lifted, "Framed one"),
            paragraph(&lifted, "Framed two"),
        ]
        .concat(),
    ));
    assert!(!shared.iter().any(is_frame), "shared frame: {shared:?}");
}

/// docxide-pdf's `paragraph-offset-frame-immediate`, the issue's case: Word
/// sets "Dear recipient" on the line after the table, as if the framed
/// paragraph between them were absent, and hangs the frame 16.8pt (`w:y`)
/// below that line's top.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn issue_case_frame_hangs_below_the_table_and_leaves_no_line() {
    let data: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/offset_text_frame.docx");
    let runs = placed(data);
    let (recipient, _) = find(&runs, "Recipient");
    let (dear, _) = find(&runs, "Dear recipient");
    let (about, _) = find(&runs, "About providing information");
    // Word: Recipient 69.78, Dear recipient 85.78; the one-line table row is
    // the only thing between them.
    assert!(
        (dear.baseline_pt - recipient.baseline_pt - 16.0).abs() < 0.3,
        "Dear recipient {} after Recipient {}",
        dear.baseline_pt,
        recipient.baseline_pt
    );
    // Word: About providing information 102.53, 16.75 below Dear recipient.
    assert_close(
        about.baseline_pt,
        dear.baseline_pt + 16.8,
        "framed paragraph",
    );
}

/// A one-row table as wide as the A4 text area, with no cell margins.
fn full_width_table(text: &str) -> String {
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9026"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9026" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#,
        paragraph("", text)
    )
}

/// An empty paragraph on an exact 12pt line in the frame `frame_xml`;
/// `extra` adds to its properties.
fn empty_framed(frame_xml: &str, extra: &str) -> String {
    format!(
        r#"<w:p><w:pPr>{frame_xml}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/>{extra}</w:pPr></w:p>"#
    )
}

/// The frame docxide-pdf's `empty-margin-frame-*` cases state: no
/// `w:vAnchor`, so margin-anchored, 3000 twips wide and 1000 twips right of
/// the text area's left edge.
const TOP_MARGIN_FRAME: &str =
    r#"<w:framePr w:w="3000" w:wrap="auto" w:hAnchor="text" w:x="1000"/>"#;

/// Empty paragraphs sharing a frame at the top margin keep the frame's box
/// there in Word 2013's layout, one line per paragraph that shows a mark,
/// though they take no flow lines (issue #2231). Word's exports of
/// docxide-pdf's `empty-margin-frame-*` cases decide each condition: the
/// compatibility-mode-14 exports keep no box, nor do hidden marks or a frame
/// beside the text area, and the paragraph ending the section adds no line.
#[test]
fn empty_paragraphs_in_a_top_margin_frame_keep_its_box_in_word_2013_layout() {
    let empty: String = empty_framed(TOP_MARGIN_FRAME, "");
    let section_end: String = empty_framed(
        TOP_MARGIN_FRAME,
        &format!("<w:sectPr>{PAGE_SETUP}</w:sectPr>"),
    );
    let body: String = format!(
        "{}{empty}{empty}{section_end}{}",
        full_width_table("Table text"),
        paragraph("", "After")
    );
    let blocks: Vec<Block> = body_blocks(&compat_frame_docx(&body, 15));
    let frames: Vec<&TextFrame> = blocks
        .iter()
        .filter_map(|block| match block {
            Block::TextFrame(frame) => Some(frame.as_ref()),
            _ => None,
        })
        .collect();
    let [frame] = frames.as_slice() else {
        panic!("expected one frame box: {blocks:?}");
    };
    assert_eq!(frame.vertical_anchor, FrameAnchor::Margin);
    assert!((frame.y).abs() < 1e-9, "{frame:?}");
    assert!((frame.x - 50.0).abs() < 1e-9, "{frame:?}");
    assert!((frame.width - FRAME_WIDTH_PT).abs() < 1e-9, "{frame:?}");
    assert!(
        matches!(
            frame.content.as_slice(),
            [Block::Paragraph(_), Block::Paragraph(_)]
        ),
        "two lines, none for the section mark: {frame:?}"
    );

    let explicit_margin: String = empty_framed(
        r#"<w:framePr w:w="3000" w:wrap="auto" w:vAnchor="margin" w:hAnchor="text" w:x="1000"/>"#,
        "",
    );
    let has_box = |body: &str, mode: u32| {
        body_blocks(&compat_frame_docx(body, mode))
            .iter()
            .any(|block| matches!(block, Block::TextFrame(_)))
    };
    let table: String = full_width_table("Table text");
    assert!(has_box(&format!("{table}{explicit_margin}"), 15));
    assert!(!has_box(&body, 14), "compatibility mode 14 keeps no box");
    let no_settings: bool = body_blocks(&frame_docx(&body))
        .iter()
        .any(|block| matches!(block, Block::TextFrame(_)));
    assert!(!no_settings, "an undeclared mode is a legacy one");

    let no_box: [(&str, String); 6] = [
        (
            "hidden marks",
            empty_framed(TOP_MARGIN_FRAME, r#"<w:rPr><w:vanish/></w:rPr>"#).repeat(2),
        ),
        (
            "beside the text area",
            empty_framed(
                r#"<w:framePr w:w="3000" w:wrap="auto" w:hAnchor="text" w:x="9100"/>"#,
                "",
            ),
        ),
        (
            "page anchored",
            empty_framed(
                r#"<w:framePr w:w="3000" w:wrap="auto" w:vAnchor="page" w:hAnchor="text" w:x="1000"/>"#,
                "",
            ),
        ),
        (
            "offset below the margin",
            empty_framed(
                r#"<w:framePr w:w="3000" w:wrap="auto" w:hAnchor="text" w:x="1000" w:y="400"/>"#,
                "",
            ),
        ),
        (
            "aligned vertically",
            empty_framed(
                r#"<w:framePr w:w="3000" w:wrap="auto" w:hAnchor="text" w:x="1000" w:yAlign="center"/>"#,
                "",
            ),
        ),
        (
            "no width",
            empty_framed(
                r#"<w:framePr w:wrap="auto" w:hAnchor="text" w:x="1000"/>"#,
                "",
            ),
        ),
    ];
    for (case, frame_group) in no_box {
        assert!(!has_box(&format!("{table}{frame_group}"), 15), "{case}");
    }
}

/// A full-width table that starts within an empty frame's box at the top
/// margin starts below it, whether the frame's paragraphs come after the
/// table or before it: Word's exports of docxide-pdf's
/// `empty-margin-frame-15-*` cases start such a table one line lower per
/// framed paragraph. Body text there stays put, as in Word's
/// `empty-frame-control` export, and so does a table that starts below the
/// box (issue #2231).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_full_width_table_in_an_empty_frames_box_starts_below_it() {
    let empty: String = empty_framed(TOP_MARGIN_FRAME, "");
    let group: String = empty.repeat(2);
    let table: String = full_width_table("Table text");
    let after: String = paragraph("", "After");
    let baseline = |body: &str, mode: u32, text: &str| {
        let runs = placed(&compat_frame_docx(body, mode));
        find(&runs, text).0.baseline_pt
    };
    // The box is the two exact 12pt lines.
    for (case, body) in [
        ("frame after the table", format!("{table}{group}{after}")),
        ("frame before the table", format!("{group}{table}{after}")),
    ] {
        let legacy: f64 = baseline(&body, 14, "Table text");
        assert_close(baseline(&body, 15, "Table text"), legacy + 24.0, case);
        let legacy_after: f64 = baseline(&body, 14, "After");
        assert_close(baseline(&body, 15, "After"), legacy_after + 24.0, case);
    }

    let lead: String = paragraph("", "Lead text");
    let text_first: String = format!("{lead}{group}{after}");
    assert_close(
        baseline(&text_first, 15, "Lead text"),
        baseline(&text_first, 14, "Lead text"),
        "body text in the box",
    );
    // Three lines of text put the table's top below the box.
    let below: String = format!("{}{table}{group}", lead.repeat(3));
    assert_close(
        baseline(&below, 15, "Table text"),
        baseline(&below, 14, "Table text"),
        "a table below the box",
    );
}

/// A table as wide as the A4 text area with one row per paragraph.
fn full_width_rows_table(paragraphs: impl IntoIterator<Item = String>) -> String {
    let rows: String = paragraphs
        .into_iter()
        .map(|paragraph| {
            format!(
                r#"<w:tr><w:tc><w:tcPr><w:tcW w:w="9026" w:type="dxa"/></w:tcPr>{paragraph}</w:tc></w:tr>"#
            )
        })
        .collect();
    format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="9026" w:type="dxa"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9026"/></w:tblGrid>{rows}</w:tbl>"#
    )
}

/// The painted non-blank runs on `page`, as text and baseline.
#[cfg(not(target_arch = "wasm32"))]
fn page_baselines(data: &[u8], page: usize) -> Vec<(String, f64)> {
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    crate::render::pdf::compiled_text_runs(&source.source, page)
        .unwrap()
        .into_iter()
        .filter(|run| !run.text.trim().is_empty())
        .map(|run| (run.text.trim_end().to_string(), run.baseline_pt))
        .collect()
}

/// The rows a full-width table carries onto the page of an empty frame's box
/// at the top margin start below the box, as Word's export of docxide-pdf's
/// `empty-margin-frame-long-prefix` starts them, and what follows them moves
/// with them. The table's first page keeps its rows, and a page the table
/// does not reach keeps its body text where it was (issue #2231).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn rows_a_table_carries_onto_an_empty_frames_box_page_start_below_it() {
    let group: String = empty_framed(TOP_MARGIN_FRAME, "").repeat(2);
    let table: String =
        full_width_rows_table((1..=70).map(|row| paragraph("", &format!("Row {row}"))));
    let after: String = paragraph("", "After");
    let carried: String = format!("{table}{group}{after}");
    let legacy_carried = |page: usize| page_baselines(&compat_frame_docx(&carried, 14), page);
    let box_carried = |page: usize| page_baselines(&compat_frame_docx(&carried, 15), page);
    assert_eq!(
        box_carried(0),
        legacy_carried(0),
        "the table's first page keeps its rows"
    );
    let (legacy, boxed) = (legacy_carried(1), box_carried(1));
    assert_eq!(
        legacy.iter().map(|(text, _)| text).collect::<Vec<_>>(),
        boxed.iter().map(|(text, _)| text).collect::<Vec<_>>(),
        "the same rows reach page 2"
    );
    assert!(
        legacy
            .first()
            .is_some_and(|(text, _)| text.starts_with("Row ")),
        "the table carries rows onto page 2: {legacy:?}"
    );
    // The box is the two exact 12pt lines.
    for ((text, legacy_pt), (_, box_pt)) in legacy.iter().zip(&boxed) {
        assert_close(*box_pt, legacy_pt + 24.0, text);
    }

    // The table ends on page 2 and body text carries the frame group onto
    // page 3, which the table does not reach.
    let filler: String = (1..=60)
        .map(|line| paragraph("", &format!("Line {line}")))
        .collect();
    let beyond: String = format!("{table}{filler}{group}{after}");
    let legacy_beyond: Vec<(String, f64)> = page_baselines(&compat_frame_docx(&beyond, 14), 2);
    assert!(
        legacy_beyond
            .first()
            .is_some_and(|(text, _)| text.starts_with("Line ")),
        "the frame group lands on page 3 below body text: {legacy_beyond:?}"
    );
    assert_eq!(
        page_baselines(&compat_frame_docx(&beyond, 15), 2),
        legacy_beyond,
        "body text on a page the table does not reach"
    );
}

/// A box that would push the rows a table carries onto its page past the
/// page's end leaves them in place, rather than move them, the group and the
/// box itself on to the next page, where the box would no longer meet them.
/// On exact 14pt rows, 49 fill the 697.9pt text area, and the two 12pt lines
/// of the box leave room for 48 below it (issue #2231).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn rows_an_empty_frames_box_would_push_off_its_page_keep_their_place() {
    let group: String = empty_framed(TOP_MARGIN_FRAME, "").repeat(2);
    let after: String = paragraph("", "After");
    let row = |row: usize| {
        format!(
            r#"<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="280" w:lineRule="exact"/></w:pPr>{}</w:p>"#,
            run(&format!("<w:t>Row {row}</w:t>"))
        )
    };
    let body = |rows: usize| {
        format!(
            "{}{group}{after}",
            full_width_rows_table((1..=rows).map(row))
        )
    };
    // Every page's painted runs, as text and baseline.
    let pages = |rows: usize, mode: u32| {
        let (document, _) = DocxParser
            .parse(
                &compat_frame_docx(&body(rows), mode),
                &ConvertOptions::default(),
            )
            .unwrap();
        let source = crate::render::typst_gen::generate_typst(&document).unwrap();
        (0..)
            .map_while(|page| crate::render::pdf::compiled_text_runs(&source.source, page).ok())
            .map(|runs| {
                runs.into_iter()
                    .filter(|run| !run.text.trim().is_empty())
                    .map(|run| (run.text.trim_end().to_string(), run.baseline_pt))
                    .collect::<Vec<(String, f64)>>()
            })
            .collect::<Vec<_>>()
    };
    let text = |runs: Option<&(String, f64)>| runs.map(|(text, _)| text.clone());

    let (legacy, boxed) = (pages(97, 14), pages(97, 15));
    assert_eq!(text(legacy[1].first()), Some("Row 50".to_string()));
    assert_eq!(text(legacy[1].last()), Some("After".to_string()));
    let carried: Vec<&(String, f64)> = legacy[1]
        .iter()
        .filter(|(text, _)| text.starts_with("Row "))
        .collect();
    assert_eq!(carried.len(), 48);
    for ((text, legacy_pt), (box_text, box_pt)) in carried.into_iter().zip(&boxed[1]) {
        assert_eq!(text, box_text);
        assert_close(*box_pt, legacy_pt + 24.0, text);
    }
    assert_eq!(
        text(boxed[1].last()),
        Some("Row 97".to_string()),
        "48 rows still fit below the box"
    );
    assert_eq!(
        text(boxed[2].first()),
        Some("After".to_string()),
        "the paragraph after them moves on with them"
    );

    assert_eq!(
        pages(98, 15),
        pages(98, 14),
        "49 rows keep their place, and nothing moves to a new page"
    );
}

/// A frame `w:x="{x_twips}"` from the page's left edge and `w:y="{y_twips}"`
/// below `v_anchor`, which keeps body text off its sides (`wrap` is
/// `notBeside` or `none`); `width` states its `w:w`, or nothing for an auto
/// width.
fn page_frame(wrap: &str, v_anchor: &str, x_twips: i32, y_twips: i32, width: &str) -> String {
    format!(
        r#"<w:framePr {width} w:wrap="{wrap}" w:vAnchor="{v_anchor}" w:hAnchor="page" w:x="{x_twips}" w:y="{y_twips}"/>"#
    )
}

/// The painted non-blank runs on `page`.
#[cfg(not(target_arch = "wasm32"))]
fn page_runs(data: &[u8], page: usize) -> Vec<crate::render::pdf::PlacedTextRun> {
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    crate::render::pdf::compiled_text_runs(&source.source, page)
        .unwrap()
        .into_iter()
        .filter(|run| !run.text.trim().is_empty())
        .collect()
}

#[cfg(not(target_arch = "wasm32"))]
fn run_named<'a>(
    runs: &'a [crate::render::pdf::PlacedTextRun],
    text: &str,
) -> &'a crate::render::pdf::PlacedTextRun {
    runs.iter()
        .find(|run| run.text.trim_end() == text)
        .unwrap_or_else(|| panic!("no run {text:?} in {runs:?}"))
}

/// Where a 12pt single-spaced line of these tests sits: its baseline below
/// the top of its line, and the pitch from one such line to the next.
#[cfg(not(target_arch = "wasm32"))]
fn line_geometry() -> (f64, f64) {
    let runs = page_runs(
        &frame_docx(&[paragraph("", "First"), paragraph("", "Second")].concat()),
        0,
    );
    let first: f64 = run_named(&runs, "First").baseline_pt;
    (
        first - MARGIN_PT,
        run_named(&runs, "Second").baseline_pt - first,
    )
}

/// A paragraph group Word frames at a fixed place on its page, keeping body
/// text off its sides (`w:wrap` `notBeside` or `none`, `w:vAnchor` `page`
/// or `margin`), is drawn at that place, not in the flow. A body line that
/// meets the frame's band starts below it; one above or below the band
/// keeps its place. Word's export of docxide-pdf's scraped
/// `massachusetts_community_sanitation` draws its letterhead frames this
/// way and starts the letter below them (issue #2231).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_frame_that_keeps_text_off_its_sides_is_drawn_at_its_page_place() {
    let (ascent, pitch) = line_geometry();
    let body_after = |frame: &str, page_break: bool| {
        let lead: String = if page_break {
            [
                paragraph("", "Lead"),
                r#"<w:p><w:r><w:br w:type="page"/></w:r></w:p>"#.to_string(),
            ]
            .concat()
        } else {
            String::new()
        };
        frame_docx(
            &[
                lead,
                paragraph(frame, "Framed one"),
                paragraph(frame, "Framed two"),
                paragraph("", "Body one"),
                paragraph("", "Body two"),
            ]
            .concat(),
        )
    };
    let width: &str = r#"w:w="3000""#;
    // (wrap, vAnchor, w:y, frame top, the first body line's top, the
    // second's): a two-line band from 60pt meets the first body line at the
    // top margin, one from 92pt only the second, and one from 10pt neither.
    let cases: [(&str, &str, i32, f64, f64, f64); 3] = [
        (
            "notBeside",
            "page",
            1200,
            60.0,
            60.0 + 2.0 * pitch,
            60.0 + 3.0 * pitch,
        ),
        (
            "none",
            "margin",
            400,
            MARGIN_PT + 20.0,
            MARGIN_PT,
            MARGIN_PT + 20.0 + 2.0 * pitch,
        ),
        ("notBeside", "page", 200, 10.0, MARGIN_PT, MARGIN_PT + pitch),
    ];
    for (wrap, v_anchor, y_twips, frame_top, body_one_top, body_two_top) in cases {
        let frame: String = page_frame(wrap, v_anchor, 2000, y_twips, width);
        for page_break in [false, true] {
            let runs = page_runs(&body_after(&frame, page_break), usize::from(page_break));
            let context = format!("wrap={wrap} vAnchor={v_anchor} y={y_twips} page {page_break}");
            for (line, text) in ["Framed one", "Framed two"].into_iter().enumerate() {
                let framed = run_named(&runs, text);
                assert_close(framed.left_pt, 100.0, &context);
                assert_close(
                    framed.baseline_pt,
                    frame_top + ascent + line as f64 * pitch,
                    &context,
                );
            }
            for (text, top) in [("Body one", body_one_top), ("Body two", body_two_top)] {
                let body = run_named(&runs, text);
                assert_close(body.left_pt, MARGIN_PT, &context);
                assert_close(body.baseline_pt, top + ascent, &format!("{context} {text}"));
            }
        }
    }
}

/// Adjacent paragraphs share a frame when their `w:framePr` agree once each
/// takes the attributes it does not state from its paragraph style's, as
/// docxide-pdf's `parse_frame_props` resolves them. The seven lines below
/// the first of `massachusetts_community_sanitation`'s letterhead state no
/// `w:hAnchor` or `w:hSpace` of their own; Word draws them in the first
/// line's frame (issue #2231).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_paragraph_takes_the_frame_attributes_it_omits_from_its_style() {
    let (ascent, pitch) = line_geometry();
    let styles: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Letterhead"><w:name w:val="Letterhead"/><w:pPr><w:framePr w:w="3000" w:hSpace="187" w:wrap="notBeside" w:vAnchor="text" w:hAnchor="page" w:x="4000" w:y="1"/></w:pPr></w:style></w:styles>"#;
    let own: &str = r#"<w:framePr w:w="3000" w:hSpace="187" w:wrap="notBeside" w:vAnchor="page" w:hAnchor="page" w:x="2000" w:y="1200"/>"#;
    let partial: &str = r#"<w:pStyle w:val="Letterhead"/><w:framePr w:w="3000" w:wrap="notBeside" w:vAnchor="page" w:x="2000" w:y="1200"/>"#;
    let document: String = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{}<w:sectPr>{PAGE_SETUP}</w:sectPr></w:body></w:document>"#,
        [
            paragraph(own, "Head line"),
            paragraph(partial, "Second line"),
            paragraph(partial, "Third line"),
            paragraph("", "Body text"),
        ]
        .concat()
    );
    let runs = page_runs(&build_docx_with_styles_xml(&document, styles), 0);
    for (line, text) in ["Head line", "Second line", "Third line"]
        .into_iter()
        .enumerate()
    {
        let framed = run_named(&runs, text);
        assert_close(framed.left_pt, 100.0, text);
        assert_close(
            framed.baseline_pt,
            60.0 + ascent + line as f64 * pitch,
            text,
        );
    }
    assert_close(
        run_named(&runs, "Body text").baseline_pt,
        60.0 + 3.0 * pitch + ascent,
        "body text below the frame",
    );
}

/// A frame without `w:w` starts at its `w:x`, and its band is as tall as its
/// content; a table whose first cell paragraph carries the frame is drawn in
/// it, its band as tall as its rows. `massachusetts_community_sanitation`'s
/// seal and governor table are such frames (issue #2231).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn an_auto_width_frame_starts_at_its_place_and_can_hold_a_table() {
    let (ascent, pitch) = line_geometry();
    let auto: String = page_frame("notBeside", "page", 2000, 1200, "");
    let runs = page_runs(
        &frame_docx(
            &[
                paragraph(&auto, "A framed line"),
                paragraph(&auto, "Another"),
                paragraph("", "Body text"),
            ]
            .concat(),
        ),
        0,
    );
    for (line, text) in ["A framed line", "Another"].into_iter().enumerate() {
        let framed = run_named(&runs, text);
        assert_close(framed.left_pt, 100.0, text);
        assert_close(
            framed.baseline_pt,
            60.0 + ascent + line as f64 * pitch,
            text,
        );
    }
    assert_close(
        run_named(&runs, "Body text").baseline_pt,
        60.0 + 2.0 * pitch + ascent,
        "body text below the frame",
    );

    // The same table in the flow, at the text area's top left, shows where
    // its rows sit in it and where the line after it starts.
    let framed_table = |frame: &str| {
        let rows: String = ["Row one", "Row two"]
            .into_iter()
            .map(|text| {
                format!(
                    r#"<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr>"#,
                    paragraph(frame, text)
                )
            })
            .collect();
        let table: String = format!(
            r#"<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>{rows}</w:tbl>"#
        );
        page_runs(
            &frame_docx(&[table, paragraph("", "Body text")].concat()),
            0,
        )
    };
    let in_flow = framed_table("");
    // From 60pt, the table's band meets the top margin's first line.
    let framed = framed_table(&page_frame("notBeside", "page", 1000, 1200, ""));
    for text in ["Row one", "Row two"] {
        let (row, flow_row) = (run_named(&framed, text), run_named(&in_flow, text));
        assert_close(row.left_pt, 50.0 + flow_row.left_pt - MARGIN_PT, text);
        assert_close(
            row.baseline_pt,
            60.0 + flow_row.baseline_pt - MARGIN_PT,
            text,
        );
    }
    assert_close(
        run_named(&framed, "Body text").baseline_pt,
        60.0 + run_named(&in_flow, "Body text").baseline_pt - MARGIN_PT,
        "body text below the table's frame",
    );
}

/// Body text clears every band it meets, whichever frame comes first in the
/// document: a line moved below one frame's band can meet another's.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn body_text_clears_overlapping_frame_bands_in_any_order() {
    let (ascent, pitch) = line_geometry();
    let width: &str = r#"w:w="3000""#;
    let lower: String = page_frame("notBeside", "page", 2000, 2000, width);
    let upper: String = page_frame("notBeside", "page", 6000, 1200, width);
    let lower_group: String = ["Lower one", "Lower two"]
        .map(|text| paragraph(&lower, text))
        .concat();
    let upper_group: String = ["Upper one", "Upper two", "Upper three"]
        .map(|text| paragraph(&upper, text))
        .concat();
    let body: String = paragraph("", "Body text");
    // The upper band, 60pt to three lines lower, reaches into the lower one,
    // which starts at 100pt.
    for (case, frames) in [
        ("upper first", format!("{upper_group}{lower_group}")),
        ("lower first", format!("{lower_group}{upper_group}")),
    ] {
        let runs = page_runs(&frame_docx(&format!("{frames}{body}")), 0);
        assert_close(
            run_named(&runs, "Upper one").baseline_pt,
            60.0 + ascent,
            case,
        );
        assert_close(
            run_named(&runs, "Lower one").baseline_pt,
            100.0 + ascent,
            case,
        );
        assert_close(
            run_named(&runs, "Body text").baseline_pt,
            100.0 + 2.0 * pitch + ascent,
            case,
        );
    }
}

/// A frame at a fixed place holds paragraphs, tables and pictures. When one
/// of its paragraphs holds something a placed frame cannot, such as a page
/// break, the whole frame stays in the flow rather than splitting its
/// paragraphs between the frame and the flow (issue #2231).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_frame_holding_what_a_placed_frame_cannot_stays_whole_in_the_flow() {
    let (ascent, pitch) = line_geometry();
    let frame: String = page_frame("notBeside", "page", 2000, 1200, r#"w:w="3000""#);
    let page_break: String = format!(
        r#"<w:p><w:pPr>{frame}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>{}{}</w:p>"#,
        run("<w:t>Framed two</w:t>"),
        run(r#"<w:br w:type="page"/><w:t>Framed three</w:t>"#)
    );
    // A break in the frame's second paragraph: its first stays put too.
    let runs = page_runs(
        &frame_docx(&[paragraph(&frame, "Framed one"), page_break.clone()].concat()),
        0,
    );
    for (line, text) in ["Framed one", "Framed two"].into_iter().enumerate() {
        let framed = run_named(&runs, text);
        assert_close(framed.left_pt, MARGIN_PT, text);
        assert_close(
            framed.baseline_pt,
            MARGIN_PT + ascent + line as f64 * pitch,
            text,
        );
    }
    // A break in its first paragraph: the paragraph after it follows the
    // break onto page 2.
    let data: Vec<u8> = frame_docx(&[page_break, paragraph(&frame, "Framed one")].concat());
    let first_page = page_runs(&data, 0);
    assert_close(
        run_named(&first_page, "Framed two").left_pt,
        MARGIN_PT,
        "Framed two",
    );
    assert!(
        first_page
            .iter()
            .all(|run| run.text.trim_end() != "Framed one"),
        "{first_page:?}"
    );
    let second_page = page_runs(&data, 1);
    for (line, text) in ["Framed three", "Framed one"].into_iter().enumerate() {
        let framed = run_named(&second_page, text);
        assert_close(framed.left_pt, MARGIN_PT, text);
        assert_close(
            framed.baseline_pt,
            MARGIN_PT + ascent + line as f64 * pitch,
            text,
        );
    }
}
