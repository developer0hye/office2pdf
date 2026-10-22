//! Word formats a list label with its paragraph mark's run properties
//! (`w:pPr/w:rPr`), resolved over the paragraph style and `w:docDefaults`,
//! with the numbering level's own `w:rPr` on top. The item's runs do not
//! format it (issue #2337).

use super::*;

const SERIF: &str = "Libertinus Serif";
const MONO: &str = "DejaVu Sans Mono";

fn fonts(family: &str) -> docx_rs::RunFonts {
    docx_rs::RunFonts::new()
        .ascii(family)
        .hi_ansi(family)
        .cs(family)
}

fn run(text: &str, family: &str, half_points: Option<usize>) -> docx_rs::Run {
    let run = docx_rs::Run::new().add_text(text).fonts(fonts(family));
    match half_points {
        Some(size) => run.size(size),
        None => run,
    }
}

/// A package with 72pt side margins, 36pt default tab stops and
/// `w:docDefaults` of 11pt `MONO`, with one numbering definition of `level`
/// and the paragraph `styles`.
fn list_docx(
    level: docx_rs::Level,
    styles: Vec<docx_rs::Style>,
    paragraphs: Vec<docx_rs::Paragraph>,
) -> Vec<u8> {
    let numberings = docx_rs::Numberings::new()
        .add_abstract_numbering(docx_rs::AbstractNumbering::new(0).add_level(level))
        .add_numbering(docx_rs::Numbering::new(1, 0));
    let mut docx = docx_rs::Docx::new()
        .page_margin(docx_rs::PageMargin::new().left(1440).right(1440))
        .default_tab_stop(720)
        .default_fonts(fonts(MONO))
        .default_size(22)
        .numberings(numberings);
    for style in styles {
        docx = docx.add_style(style);
    }
    for paragraph in paragraphs {
        docx = docx.add_paragraph(paragraph);
    }
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    cursor.into_inner()
}

/// `(%1)` with a hanging indent and no run properties of its own.
fn plain_level() -> docx_rs::Level {
    docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("(%1)"),
        docx_rs::LevelJc::new("left"),
    )
    .indent(
        Some(720),
        Some(docx_rs::SpecialIndentType::Hanging(360)),
        None,
        None,
    )
}

fn item(runs: Vec<docx_rs::Run>) -> docx_rs::Paragraph {
    runs.into_iter().fold(
        docx_rs::Paragraph::new()
            .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
        docx_rs::Paragraph::add_run,
    )
}

/// The family and size of the label painted before `body` on its baseline.
#[cfg(not(target_arch = "wasm32"))]
fn painted_label_face(data: &[u8], label: &str, body: &str) -> (String, f64) {
    let (document, _warnings) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
    let body_run = runs
        .iter()
        .find(|run| run.text.contains(body))
        .unwrap_or_else(|| panic!("{body:?} is not painted: {runs:?}"));
    let label_run = runs
        .iter()
        .find(|run| {
            run.text.trim() == label
                && (run.baseline_pt - body_run.baseline_pt).abs() < 0.5
                && run.left_pt < body_run.left_pt
        })
        .unwrap_or_else(|| panic!("no {label:?} before {body:?}: {runs:?}"));
    (label_run.family.clone(), label_run.font_size_pt)
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_list_label_takes_its_face_and_size_from_the_paragraph_mark() {
    let label_style = docx_rs::Style::new("LabelStyle", docx_rs::StyleType::Paragraph)
        .name("Label Style")
        .fonts(fonts(SERIF))
        .size(28);
    let cases: [(&str, docx_rs::Level, docx_rs::Paragraph, (&str, f64)); 5] = [
        (
            // docxide-pdf's `k4_regulations_synthesis_document` `(3)`: the
            // runs disagree on size, the mark states it.
            "mark states what the runs disagree on",
            plain_level(),
            item(vec![
                run("Body of the item ", SERIF, Some(20)),
                run("B", SERIF, None).bold(),
            ])
            .fonts(fonts(SERIF))
            .size(20),
            (SERIF, 10.0),
        ),
        (
            // Its `5.`: the runs agree, but no mark property means the
            // document defaults.
            "mark without properties takes the document defaults",
            plain_level(),
            item(vec![run("Body of the item", SERIF, Some(20))]),
            (MONO, 11.0),
        ),
        (
            "mark properties over runs that agree",
            plain_level(),
            item(vec![run("Body of the item", MONO, Some(18))])
                .fonts(fonts(SERIF))
                .size(26),
            (SERIF, 13.0),
        ),
        (
            "paragraph style under a mark without properties",
            plain_level(),
            item(vec![run("Body of the item", MONO, Some(18))]).style("LabelStyle"),
            (SERIF, 14.0),
        ),
        (
            // Its `a.` items: the level's own `w:rPr` outranks the mark.
            "level run properties over the mark",
            plain_level().fonts(fonts(SERIF)).size(18),
            item(vec![run("Body of the item", MONO, Some(24))])
                .fonts(fonts(MONO))
                .size(24),
            (SERIF, 9.0),
        ),
    ];
    for (case, level, paragraph, (family, size)) in cases {
        let data = list_docx(level, vec![label_style.clone()], vec![paragraph]);
        let (painted_family, painted_size) = painted_label_face(&data, "(1)", "Body of the item");
        assert_eq!(painted_family, family, "{case}");
        assert!(
            (painted_size - size).abs() < 0.01,
            "{case}: label painted at {painted_size}pt, expected {size}pt"
        );
    }
}

/// Word puts a number inline when its level indents by zero; it is still the
/// paragraph mark that formats it (issue #2337).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_inline_list_number_takes_its_face_and_size_from_the_paragraph_mark() {
    let flush_level = docx_rs::Level::new(
        0,
        docx_rs::Start::new(1),
        docx_rs::NumberFormat::new("decimal"),
        docx_rs::LevelText::new("(%1)"),
        docx_rs::LevelJc::new("left"),
    )
    .indent(
        Some(0),
        Some(docx_rs::SpecialIndentType::Hanging(0)),
        None,
        None,
    );
    for (mark_family, mark_size, runs_family, runs_size) in
        [(SERIF, 20, MONO, 18), (MONO, 26, SERIF, 20)]
    {
        let data = list_docx(
            flush_level.clone(),
            Vec::new(),
            vec![
                item(vec![run("Body of the item", runs_family, Some(runs_size))])
                    .fonts(fonts(mark_family))
                    .size(mark_size),
            ],
        );
        let (painted_family, painted_size) = painted_label_face(&data, "(1)", "Body of the item");
        assert_eq!(painted_family, mark_family);
        assert!(
            (painted_size - mark_size as f64 / 2.0).abs() < 0.01,
            "label painted at {painted_size}pt under a {mark_size} half-point mark"
        );
    }
}

/// Word ignores the mark's underline and vertical position for the label, as
/// LibreOffice's `IsIgnoredCharFormatForNumbering` records; its weight,
/// posture and colour carry (issue #2337).
#[test]
fn docx_list_label_style_drops_the_marks_underline_and_vertical_position() {
    let data = list_docx(
        plain_level(),
        Vec::new(),
        vec![
            item(vec![run("Body of the item", MONO, Some(20))]).run_property(
                docx_rs::RunProperty::new()
                    .fonts(fonts(SERIF))
                    .size(20)
                    .bold()
                    .italic()
                    .color("C00000")
                    .underline("single")
                    .vert_align(docx_rs::VertAlignType::SuperScript),
            ),
        ],
    );
    let (document, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(flow) = &document.pages[0] else {
        panic!("expected a flow page");
    };
    let Some(Block::List(list)) = flow.content.first() else {
        panic!("expected a list: {:?}", flow.content);
    };
    let label = list.items[0].content[0]
        .style
        .numbering_label_style
        .as_deref()
        .expect("a numbered paragraph carries its label's formatting");
    assert_eq!(label.font_family.as_deref(), Some(SERIF));
    assert_eq!(label.font_size, Some(10.0));
    assert_eq!(label.bold, Some(true));
    assert_eq!(label.italic, Some(true));
    assert_eq!(
        label.color,
        Some(Color {
            r: 0xC0,
            g: 0,
            b: 0
        })
    );
    assert_eq!(label.underline, None);
    assert_eq!(label.vertical_align, None);
}

/// Typst gives a list level one numbering function, but each Word item's
/// label still follows its own mark when the marks disagree, as in
/// docxide-pdf's `case73`, whose `I.` is 15pt beside 11pt labels elsewhere
/// (issue #2337). Only a label wider than the 18pt hanging indent moves its
/// own body on to the next tab stop.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_list_labels_follow_their_own_marks_within_one_level() {
    let marks: [(&str, usize); 4] = [(SERIF, 26), (SERIF, 20), (MONO, 24), (SERIF, 60)];
    let bodies = ["First body", "Second body", "Third body", "Fourth body"];
    let data = list_docx(
        plain_level(),
        Vec::new(),
        marks
            .iter()
            .zip(bodies)
            .map(|((family, size), body)| {
                item(vec![run(body, MONO, Some(20))])
                    .fonts(fonts(family))
                    .size(*size)
            })
            .collect(),
    );
    let (document, _warnings) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let source = crate::render::typst_gen::generate_typst(&document).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&source.source, 0).unwrap();
    for (index, ((family, size), body)) in marks.iter().zip(bodies).enumerate() {
        let label = format!("({})", index + 1);
        let (painted_family, painted_size) = painted_label_face(&data, &label, body);
        assert_eq!(painted_family, *family, "{label}");
        assert!(
            (painted_size - *size as f64 / 2.0).abs() < 0.01,
            "{label} painted at {painted_size}pt under a {size} half-point mark"
        );
        let body_run = runs.iter().find(|run| run.text.contains(body)).unwrap();
        // The level hangs 18pt from a 36pt indent past the 72pt margin. The
        // 12pt monospaced `(3)` is about 21.7pt wide, so it and the 30pt `(4)`
        // overflow it; the document-default label every item shared before
        // would overflow it too.
        let expected_left: f64 = [108.0, 108.0, 144.0, 144.0][index];
        assert!(
            (body_run.left_pt - expected_left).abs() < 0.01,
            "{body} starts at {}pt after {label}, expected {expected_left}pt",
            body_run.left_pt
        );
    }
}

/// A list that restarts onto a number it already used, here through a second
/// `w:num` with a `w:startOverride`, gives the numbering function two labels
/// it cannot tell apart. They keep what their marks share; every other label
/// keeps its own (issue #2337).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_list_labels_sharing_a_restarted_number_keep_what_their_marks_share() {
    let numberings = docx_rs::Numberings::new()
        .add_abstract_numbering(docx_rs::AbstractNumbering::new(0).add_level(plain_level()))
        .add_numbering(docx_rs::Numbering::new(1, 0))
        .add_numbering(
            docx_rs::Numbering::new(2, 0).add_override(docx_rs::LevelOverride::new(0).start(1)),
        );
    let items: [(usize, usize, &str); 3] = [
        (1, 30, "First body"),
        (1, 20, "Second body"),
        (2, 20, "Restarted body"),
    ];
    let mut docx = docx_rs::Docx::new()
        .page_margin(docx_rs::PageMargin::new().left(1440).right(1440))
        .default_fonts(fonts(MONO))
        .default_size(22)
        .numberings(numberings);
    for (num_id, mark_size, body) in items {
        docx = docx.add_paragraph(
            docx_rs::Paragraph::new()
                .add_run(run(body, MONO, Some(20)))
                .numbering(
                    docx_rs::NumberingId::new(num_id),
                    docx_rs::IndentLevel::new(0),
                )
                .fonts(fonts(SERIF))
                .size(mark_size),
        );
    }
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let first = painted_label_face(&data, "(1)", "First body");
    let second = painted_label_face(&data, "(2)", "Second body");
    let restarted = painted_label_face(&data, "(1)", "Restarted body");
    assert_eq!(second, (SERIF.to_string(), 10.0), "(2) keeps its own mark");
    assert_eq!(
        first, restarted,
        "both `(1)` labels take the formatting their 15pt and 10pt marks share"
    );
    assert_eq!(first.0, SERIF, "the marks share their face");
    assert!(
        (restarted.1 - 15.0).abs() > 0.01,
        "the restarted label does not take the first mark's 15pt: {restarted:?}"
    );
}
