//! Tabs in a Word body paragraph: stops count from the margin, a tab that
//! runs past the line moves to the next line with the word after it (issue
//! #2226), and a leader is drawn in the tab's own run formatting (issue
//! #2244).

use super::*;

const MARGIN_PT: f64 = 72.0;
/// A4's 11906 twips less two 1440-twip margins.
const TEXT_WIDTH_PT: f64 = 451.3;
/// The centre stop at 5018 twips and the left stop at 8625 twips, as in
/// docxide-pdf's `leading-tab-*` cases.
const ISSUE_STOPS: &str =
    r#"<w:tabs><w:tab w:val="center" w:pos="5018"/><w:tab w:val="left" w:pos="8625"/></w:tabs>"#;
const ISSUE_RIGHT_STOPS: &str =
    r#"<w:tabs><w:tab w:val="center" w:pos="5018"/><w:tab w:val="right" w:pos="8625"/></w:tabs>"#;
const CENTRE_STOP_PT: f64 = 250.9;
const GREETING: &str = "Public greeting text";

fn run(half_points: u32, content: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="{half_points}"/></w:rPr>{content}</w:r>"#
    )
}

fn tabs(half_points: &[u32]) -> String {
    half_points
        .iter()
        .map(|size| run(*size, "<w:tab/>"))
        .collect()
}

fn paragraph(paragraph_properties: &str, runs: &str) -> String {
    // `w:pPr` children in schema order: tabs, spacing, ind.
    let (tab_stops, indent) = match paragraph_properties.find("<w:ind") {
        Some(at) => paragraph_properties.split_at(at),
        None => (paragraph_properties, ""),
    };
    format!(
        r#"<w:p><w:pPr>{tab_stops}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>{indent}</w:pPr>{runs}</w:p>"#
    )
}

/// An A4 document with 72pt margins and a 708-twip default stop holding
/// `paragraphs`, then a `Public body` paragraph.
fn tab_docx(paragraphs: &[String]) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{}{}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>"#,
        paragraphs.concat(),
        paragraph("", &run(28, "<w:t>Public body</w:t>"))
    );
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    let parts: [(&str, &str); 5] = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        (
            "word/_rels/document.xml.rels",
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>"#,
        ),
        (
            "word/settings.xml",
            r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="708"/></w:settings>"#,
        ),
        ("word/document.xml", &document),
    ];
    for (name, content) in parts {
        zip.start_file(name, options).unwrap();
        std::io::Write::write_all(&mut zip, content.as_bytes()).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

/// The painted text runs, without the zero-width strut a tab leaves on its
/// line, and each run's advance width.
#[cfg(not(target_arch = "wasm32"))]
fn placed(data: &[u8]) -> Vec<(crate::render::pdf::PlacedTextRun, f64)> {
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    let text = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
    let glyphs = crate::render::pdf::compiled_glyph_runs(&source.source, 0, true).unwrap();
    text.into_iter()
        .filter(|run| !run.text.trim_matches('\u{2060}').is_empty())
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
        .find(|(run, _)| run.text == text)
        .unwrap_or_else(|| panic!("no run {text:?} in {runs:?}"))
}

#[cfg(not(target_arch = "wasm32"))]
fn assert_close(actual: f64, expected: f64, context: &str) {
    assert!(
        (actual - expected).abs() < 0.05,
        "{context}: {actual} != {expected}"
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn word_paragraph_tab_stops_count_from_the_margin_not_the_indent() {
    // A first-line indent: the tab still centres the text on the stop.
    let data = tab_docx(&[paragraph(
        &format!(r#"{ISSUE_STOPS}<w:ind w:firstLine="539"/>"#),
        &(tabs(&[28]) + &run(28, &format!("<w:t>{GREETING}</w:t>"))),
    )]);
    let runs = placed(&data);
    let (text, width) = find(&runs, GREETING);
    assert_close(
        text.left_pt,
        MARGIN_PT + CENTRE_STOP_PT - width / 2.0,
        "first-line indent, centre stop",
    );

    // A left indent: a left stop at 144pt lands 144pt from the margin.
    let data = tab_docx(&[paragraph(
        r#"<w:tabs><w:tab w:val="left" w:pos="2880"/></w:tabs><w:ind w:left="1440"/>"#,
        &run(28, "<w:t>Lead</w:t><w:tab/><w:t>Next</w:t>"),
    )]);
    let runs = placed(&data);
    assert_close(
        find(&runs, "Next").0.left_pt,
        MARGIN_PT + 144.0,
        "left indent, left stop",
    );

    // A right stop on the text edge keeps an indented entry's number on its
    // line, ending at the edge.
    let data = tab_docx(&[paragraph(
        r#"<w:tabs><w:tab w:val="right" w:pos="9026"/></w:tabs><w:ind w:left="1440"/>"#,
        &run(28, "<w:t>Entry</w:t><w:tab/><w:t>12</w:t>"),
    )]);
    let runs = placed(&data);
    let (number, width) = find(&runs, "12");
    assert_close(
        number.left_pt + width,
        MARGIN_PT + TEXT_WIDTH_PT,
        "left indent, right stop at the edge",
    );
    assert_close(
        number.baseline_pt,
        find(&runs, "Entry").0.baseline_pt,
        "the number stays on the entry's line",
    );

    // Without stops, the default grid counts from the margin too: from 54pt
    // plus a short word the tab reaches 2 x 35.4pt.
    let data = tab_docx(&[paragraph(
        r#"<w:ind w:left="1080"/>"#,
        &run(28, "<w:t>A</w:t><w:tab/><w:t>B</w:t>"),
    )]);
    let runs = placed(&data);
    let (lead, lead_width) = find(&runs, "A");
    assert!(54.0 + lead_width < 70.8, "{lead_width}");
    assert_close(lead.left_pt, MARGIN_PT + 54.0, "left indent origin");
    assert_close(
        find(&runs, "B").0.left_pt,
        MARGIN_PT + 70.8,
        "left indent, default grid",
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_hanging_indent_is_the_first_line_tab_stop() {
    let numbered = |indent: &str, tab_stops: &str, runs: &str| {
        placed(&tab_docx(&[paragraph(
            &format!("{tab_stops}{indent}"),
            &run(28, runs),
        )]))
    };
    let flush_hanging = r#"<w:ind w:left="1440" w:hanging="1440"/>"#;

    // The tab skips the 35.4pt default stop for the 72pt left indent.
    let runs = numbered(flush_hanging, "", "<w:t>1.</w:t><w:tab/><w:t>Body</w:t>");
    assert!(find(&runs, "1.").1 < 35.4);
    assert_close(
        find(&runs, "Body").0.left_pt,
        MARGIN_PT + 72.0,
        "hanging indent stop",
    );

    // An explicit stop before the indent comes first.
    let runs = numbered(
        flush_hanging,
        r#"<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>"#,
        "<w:t>1.</w:t><w:tab/><w:t>Body</w:t>",
    );
    assert_close(
        find(&runs, "Body").0.left_pt,
        MARGIN_PT + 36.0,
        "explicit stop before the indent",
    );

    // A pen already past the indent takes the next default stop.
    let runs = numbered(
        flush_hanging,
        "",
        "<w:t>Long numbered label</w:t><w:tab/><w:t>Body</w:t>",
    );
    let (label, label_width) = find(&runs, "Long numbered label");
    let pen = label.left_pt - MARGIN_PT + label_width;
    assert!(pen > 72.0, "{pen}");
    assert_close(
        find(&runs, "Body").0.left_pt,
        MARGIN_PT + ((pen / 35.4).floor() + 1.0) * 35.4,
        "default stop past the indent",
    );

    // A contents entry: the title starts at the indent and its page number
    // ends at the right stop on the same line.
    let runs = numbered(
        r#"<w:ind w:left="1134" w:hanging="567"/>"#,
        r#"<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9026"/></w:tabs>"#,
        "<w:t>5</w:t><w:tab/><w:t>Vacancy in the office</w:t><w:tab/><w:t>6</w:t>",
    );
    let (title, _) = find(&runs, "Vacancy in the office");
    assert_close(title.left_pt, MARGIN_PT + 56.7, "contents title");
    let (page, page_width) = find(&runs, "6");
    assert_close(
        page.left_pt + page_width,
        MARGIN_PT + TEXT_WIDTH_PT,
        "contents page number",
    );
    assert_close(page.baseline_pt, title.baseline_pt, "one contents line");
}

/// `probe` is a paragraph whose tabs run past the line; `lines` lists the
/// sizes of the tabs Word leaves on each line above the text's own line.
/// The text's line, and every line after it, must sit where a stack of
/// one-line paragraphs of those sizes puts them: a line holding only tabs is
/// as tall as their runs.
#[cfg(not(target_arch = "wasm32"))]
fn assert_text_on_line_below(probe: &str, lines: &[&[u32]], expected_left_pt: impl Fn(f64) -> f64) {
    let runs = placed(&tab_docx(&[probe.to_string()]));
    let (text, width) = find(&runs, GREETING);
    assert_close(text.left_pt, expected_left_pt(*width), "text position");

    let mut control: Vec<String> = lines
        .iter()
        .map(|sizes| {
            let marks: String = sizes
                .iter()
                .map(|size| run(*size, "<w:t>x</w:t>"))
                .collect();
            paragraph("", &marks)
        })
        .collect();
    control.push(paragraph("", &run(28, &format!("<w:t>{GREETING}</w:t>"))));
    let control_runs = placed(&tab_docx(&control));
    for label in [GREETING, "Public body"] {
        assert_close(
            find(&runs, label).0.baseline_pt,
            find(&control_runs, label).0.baseline_pt,
            &format!("{label} baseline after {lines:?}"),
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn word_paragraph_tabs_past_the_line_move_to_the_next_line() {
    let first_line = r#"<w:ind w:firstLine="539"/>"#;
    let greeting = run(28, &format!("<w:t>{GREETING}</w:t>"));
    let centred = |width: f64| MARGIN_PT + CENTRE_STOP_PT - width / 2.0;
    let probe = |stops: &str, prefix: &str, sizes: &[u32]| {
        paragraph(
            &format!("{stops}{first_line}"),
            &(prefix.to_string() + &tabs(sizes) + &greeting),
        )
    };

    // One tab reaches the centre stop and the text fits on the first line.
    assert_text_on_line_below(&probe(ISSUE_STOPS, "", &[28]), &[], centred);
    // The second tab reaches the left stop, where the text's first word no
    // longer fits: that tab starts the next line and centres the text there.
    assert_text_on_line_below(&probe(ISSUE_STOPS, "", &[28, 28]), &[&[28]], centred);
    // The line holding only the first tab takes that tab's size, small or
    // large.
    assert_text_on_line_below(&probe(ISSUE_STOPS, "", &[20, 28]), &[&[20]], centred);
    assert_text_on_line_below(&probe(ISSUE_STOPS, "", &[40, 28]), &[&[40]], centred);
    // A third tab past the last stop would land beyond the line's end.
    assert_text_on_line_below(
        &probe(ISSUE_STOPS, "", &[20, 28, 28]),
        &[&[20, 28]],
        centred,
    );
    // Four tabs fill two lines: the second also ends where its next tab's
    // word no longer fits.
    assert_text_on_line_below(
        &probe(ISSUE_STOPS, "", &[20, 28, 28, 28]),
        &[&[20, 28], &[28]],
        centred,
    );
    // Visible text before the tabs stays on the first line.
    let prefixed = probe(ISSUE_STOPS, &run(28, "<w:t>Prefix</w:t>"), &[20, 28]);
    let runs = placed(&tab_docx(std::slice::from_ref(&prefixed)));
    assert_close(
        find(&runs, "Prefix").0.left_pt,
        MARGIN_PT + 26.95,
        "prefix after the first-line indent",
    );
    assert_text_on_line_below(&prefixed, &[&[28]], centred);
    // Right-aligned on the second stop, the text fits on the first line.
    assert_text_on_line_below(&probe(ISSUE_RIGHT_STOPS, "", &[20, 28]), &[], |width| {
        MARGIN_PT + 431.25 - width
    });
}

/// A run in `font` at `half_points`, painted `color` (hex RRGGBB).
#[cfg(not(target_arch = "wasm32"))]
fn styled_run(font: &str, half_points: u32, color: &str, content: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:color w:val="{color}"/><w:sz w:val="{half_points}"/></w:rPr>{content}</w:r>"#
    )
}

#[cfg(not(target_arch = "wasm32"))]
fn rgba(hex: &str) -> [u8; 4] {
    let channel = |at: usize| u8::from_str_radix(&hex[at..at + 2], 16).unwrap();
    [channel(0), channel(2), channel(4), 255]
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn a_tab_leader_takes_the_face_size_and_colour_of_the_tab_run() {
    // Word draws a leader as the tab character's own ink, so it takes the
    // formatting of the run that holds the `w:tab`, whatever the text around
    // it states. The leader used to take the engine's default face and size:
    // docxide-pdf's case10 rows painted 11pt Libertinus periods where Word
    // paints the tab run's 12pt Aptos (issue #2244).
    for (leader, glyph) in [("dot", '.'), ("hyphen", '-'), ("underscore", '_')] {
        // The tab run against text in the other face, size and colour.
        for (tab_font, tab_half_points, tab_color, text_font) in [
            ("Times New Roman", 28, "C00000", "Arial"),
            ("Arial", 18, "1F4E79", "Times New Roman"),
        ] {
            let text = |content: &str| {
                styled_run(text_font, 20, "000000", &format!("<w:t>{content}</w:t>"))
            };
            let row = paragraph(
                &format!(
                    r#"<w:tabs><w:tab w:val="right" w:leader="{leader}" w:pos="7920"/></w:tabs>"#
                ),
                &(text("Contents")
                    + &styled_run(tab_font, tab_half_points, tab_color, "<w:tab/>")
                    + &text("7")),
            );
            // Text in the tab run's face shows which family that face
            // resolves to on this host.
            let face = paragraph(
                "",
                &styled_run(tab_font, 20, "000000", "<w:t>Reference</w:t>"),
            );
            let runs = placed(&tab_docx(&[row, face]));
            let case = format!("{leader} leader on a {tab_font} tab");
            let family: &str = &find(&runs, "Reference").0.family;
            let contents = &find(&runs, "Contents").0;
            assert_ne!(
                family, contents.family,
                "{case}: the two faces must resolve apart to tell them apart"
            );

            let leaders: Vec<&crate::render::pdf::PlacedTextRun> = runs
                .iter()
                .map(|(run, _)| run)
                .filter(|run| run.text.chars().all(|ch| ch == glyph))
                .collect();
            assert!(
                leaders.len() > 10,
                "{case}: expected a row of leader marks, got {runs:?}"
            );
            for mark in leaders {
                assert_eq!(mark.family, family, "{case}: leader face");
                assert_close(
                    mark.font_size_pt,
                    f64::from(tab_half_points) / 2.0,
                    &format!("{case}: leader size"),
                );
                assert_eq!(
                    mark.fill_rgba,
                    Some(rgba(tab_color)),
                    "{case}: leader colour"
                );
            }
            // The text either side keeps its own run's formatting.
            assert_close(contents.font_size_pt, 10.0, &format!("{case}: text size"));
            assert_eq!(
                find(&runs, "7").0.fill_rgba,
                Some(rgba("000000")),
                "{case}: text colour"
            );
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn the_tab_leaders_fixture_draws_each_rows_dots_in_its_runs_face_and_size() {
    // docxide-pdf's case10: each row is one unformatted run that resolves
    // `Normal`'s Aptos 12pt, and Word's export draws its dots in that face
    // and size (issue #2244).
    let data: &[u8] = include_bytes!("../../../../tests/fixtures/docx/docxide/tab_leaders.docx");
    let runs = placed(data);
    for label in [
        "Introduction",
        "Background",
        "Methods",
        "Results",
        "Conclusion",
    ] {
        let row = &find(&runs, label).0;
        let dots: Vec<&crate::render::pdf::PlacedTextRun> = runs
            .iter()
            .map(|(run, _)| run)
            .filter(|run| run.text == "." && (run.baseline_pt - row.baseline_pt).abs() < 0.01)
            .collect();
        assert!(
            dots.len() > 50,
            "{label}: expected a row of dots, got {dots:?}"
        );
        for dot in dots {
            assert_eq!(dot.family, row.family, "{label}: leader face");
            assert_close(dot.font_size_pt, 12.0, &format!("{label}: leader size"));
        }
    }
}
