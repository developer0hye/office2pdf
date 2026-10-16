use super::*;
use crate::ir::{TabAlignment, TabLeader, TabStop};
use crate::render::pdf::{PlacedTextRun, compiled_text_runs_with_images};

fn compile_entry(right_indent: f64, alignment: TabAlignment, body: &str) -> Vec<PlacedTextRun> {
    compile_entry_with_suffix(right_indent, alignment, body, true)
}

fn compile_entry_with_suffix(
    right_indent: f64,
    alignment: TabAlignment,
    body: &str,
    suffix: bool,
) -> Vec<PlacedTextRun> {
    compile_indented_entry(right_indent, alignment, body, suffix, 0.0, 0.0)
}

fn compile_indented_entry(
    right_indent: f64,
    alignment: TabAlignment,
    body: &str,
    suffix: bool,
    left: f64,
    first: f64,
) -> Vec<PlacedTextRun> {
    let mut page = make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            indent_right: Some(right_indent),
            indent_left: Some(left),
            indent_first_line: Some(first),
            tab_stops: Some(vec![
                TabStop {
                    position: 54.0,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                },
                TabStop {
                    position: 360.0,
                    alignment,
                    leader: TabLeader::None,
                },
            ]),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: if suffix {
                format!("7.\t{body}\t23")
            } else {
                format!("7.\t{body}")
            },
            style: TextStyle {
                font_family: Some("Noto Serif".to_string()),
                font_size: Some(11.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
            inline_box: None,
        }],
    })]);
    let Page::Flow(flow) = &mut page else {
        unreachable!()
    };
    flow.size = PageSize {
        width: 480.0,
        height: 500.0,
    };
    flow.margins = Margins {
        top: 40.0,
        bottom: 40.0,
        left: 40.0,
        right: 40.0,
    };
    let doc = make_doc(vec![page]);
    let output = generate_typst(&doc).unwrap();
    compiled_text_runs_with_images(&output.source, &output.images, 0).unwrap()
}

fn label<'a>(runs: &'a [PlacedTextRun], text: &str) -> &'a PlacedTextRun {
    runs.iter()
        .find(|run| run.text == text)
        .unwrap_or_else(|| panic!("missing {text}: {runs:?}"))
}

#[test]
fn tab_overhang_aligned_short_entries_ignore_the_right_indent_at_the_stop() {
    for alignment in [
        TabAlignment::Right,
        TabAlignment::Center,
        TabAlignment::Decimal,
    ] {
        let control = compile_entry(0.0, alignment, "Eligibility");
        let control_number = label(&control, "23");
        for right_indent in [48.0, 72.0] {
            let runs = compile_entry(right_indent, alignment, "Eligibility");
            let body = label(&runs, "Eligibility");
            let number = label(&runs, "23");
            assert!((body.left_pt - 94.0).abs() < 0.01, "{runs:?}");
            assert!(
                (body.baseline_pt - number.baseline_pt).abs() < 0.01,
                "a reachable aligned tab may paint in the right indent: {alignment:?}, indent={right_indent}: {runs:?}"
            );
            assert!(
                (number.left_pt - control_number.left_pt).abs() < 0.01,
                "{runs:?}"
            );
        }
    }
}

#[test]
fn tab_overhang_does_not_expand_the_ordinary_text_wrap_region() {
    let body = "Eligibility and application requirements for every participating organization";
    let runs = compile_entry(100.0, TabAlignment::Right, body);
    let control = compile_entry_with_suffix(100.0, TabAlignment::Right, body, false);
    let unindented = compile_entry(0.0, TabAlignment::Right, "Eligibility");
    let number = label(&runs, "23");
    let body_runs: Vec<_> = runs
        .iter()
        .filter(|run| run.text != "7." && run.text != "23")
        .collect();
    let mut baselines: Vec<_> = body_runs.iter().map(|run| run.baseline_pt).collect();
    baselines.sort_by(f64::total_cmp);
    baselines.dedup_by(|a, b| (*a - *b).abs() < 0.01);
    assert!(
        baselines.len() >= 2,
        "the body still wraps inside the paragraph's right indent: {runs:?}"
    );
    assert!(
        (number.baseline_pt - baselines.last().unwrap()).abs() < 0.01,
        "the trailing aligned number follows the final body line: {runs:?}"
    );
    assert!(
        (number.left_pt - label(&unindented, "23").left_pt).abs() < 0.01,
        "{runs:?}"
    );
    let body_lines = |runs: &[PlacedTextRun]| {
        runs.iter()
            .filter(|run| run.text != "7." && run.text != "23")
            .map(|run| {
                (
                    run.text.clone(),
                    (run.left_pt * 100.0).round() as i64,
                    (run.baseline_pt * 100.0).round() as i64,
                )
            })
            .collect::<Vec<_>>()
    };
    assert_eq!(
        body_lines(&runs),
        body_lines(&control),
        "the aligned field must not widen the body wrap region"
    );
}

#[test]
fn tab_overhang_hanging_indents_keep_the_column_tab_origin() {
    let control = compile_entry(0.0, TabAlignment::Right, "Eligibility");
    for (left, first) in [(54.0, -36.0), (72.0, -54.0), (18.0, 12.0)] {
        let runs =
            compile_indented_entry(72.0, TabAlignment::Right, "Eligibility", true, left, first);
        assert!(
            (label(&runs, "7.").left_pt - (40.0 + left + first)).abs() < 0.01,
            "{runs:?}"
        );
        assert!(
            (label(&runs, "Eligibility").left_pt - 94.0).abs() < 0.01,
            "{runs:?}"
        );
        assert!(
            (label(&runs, "23").left_pt - label(&control, "23").left_pt).abs() < 0.01,
            "{runs:?}"
        );
        assert!(
            (label(&runs, "Eligibility").baseline_pt - label(&runs, "23").baseline_pt).abs() < 0.01,
            "{runs:?}"
        );
    }
}

#[test]
fn tab_overhang_docx_package_keeps_the_hanging_entry_on_one_line() {
    use crate::parser::{Parser, docx::DocxParser};
    let paragraph = docx_rs::Paragraph::new()
        .indent(
            Some(1080),
            Some(docx_rs::SpecialIndentType::Hanging(720)),
            Some(1440),
            None,
        )
        .add_tab(
            docx_rs::Tab::new()
                .val(docx_rs::TabValueType::Left)
                .pos(1080),
        )
        .add_tab(
            docx_rs::Tab::new()
                .val(docx_rs::TabValueType::Right)
                .pos(7200),
        )
        .add_run(
            docx_rs::Run::new()
                .add_text("7.")
                .add_tab()
                .add_text("Eligibility")
                .add_tab()
                .add_text("23"),
        );
    let package = docx_rs::Docx::new()
        .page_size(9600, 10000)
        .page_margin(
            docx_rs::PageMargin::new()
                .left(800)
                .right(800)
                .top(800)
                .bottom(800),
        )
        .add_paragraph(paragraph);
    let mut bytes = std::io::Cursor::new(Vec::new());
    package.build().pack(&mut bytes).unwrap();
    let (doc, _) = DocxParser
        .parse(bytes.get_ref(), &crate::ConvertOptions::default())
        .unwrap();
    let output = generate_typst(&doc).unwrap();
    let runs = compiled_text_runs_with_images(&output.source, &output.images, 0).unwrap();
    let first = label(&runs, "7.");
    let title = label(&runs, "Eligibility");
    let page_number = label(&runs, "23");
    assert!((first.left_pt - 58.0).abs() < 0.01, "{runs:?}");
    assert!((title.left_pt - 94.0).abs() < 0.01, "{runs:?}");
    assert!(
        (first.baseline_pt - title.baseline_pt).abs() < 0.01
            && (title.baseline_pt - page_number.baseline_pt).abs() < 0.01,
        "{runs:?}"
    );
}
