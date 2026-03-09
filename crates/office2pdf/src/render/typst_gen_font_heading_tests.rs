use super::*;

#[test]
fn test_font_substitution_calibri_produces_fallback_list() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Calibri text".to_string(),
            style: TextStyle {
                font_family: Some("Calibri".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"font: ("Calibri", "Carlito", "Liberation Sans")"#),
        "Expected font fallback list for Calibri in: {result}"
    );
}

#[test]
fn test_font_substitution_arial_produces_fallback_list() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Arial text".to_string(),
            style: TextStyle {
                font_family: Some("Arial".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"font: ("Arial", "Liberation Sans", "Arimo")"#),
        "Expected font fallback list for Arial in: {result}"
    );
}

#[test]
fn test_font_substitution_unknown_font_no_fallback() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Custom text".to_string(),
            style: TextStyle {
                font_family: Some("Helvetica".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"font: "Helvetica""#),
        "Unknown font should use simple quoted string in: {result}"
    );
    assert!(
        !result.contains("font: (\"Helvetica\""),
        "Unknown font should not use array syntax in: {result}"
    );
}

#[test]
fn test_font_substitution_times_new_roman() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "TNR text".to_string(),
            style: TextStyle {
                font_family: Some("Times New Roman".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"font: ("Times New Roman", "Liberation Serif", "Tinos")"#),
        "Expected font fallback list for Times New Roman in: {result}"
    );
}

#[test]
fn test_font_family_infers_medium_weight_from_family_name() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Title".to_string(),
            style: TextStyle {
                font_family: Some("Pretendard Medium".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"weight: "medium""#),
        "Expected medium weight inferred from family name in: {result}"
    );
}

#[test]
fn test_font_family_infers_extrabold_weight_from_family_name() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Heading".to_string(),
            style: TextStyle {
                font_family: Some("Pretendard ExtraBold".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains(r#"weight: "extrabold""#),
        "Expected extrabold weight inferred from family name in: {result}"
    );
}

#[test]
fn test_generate_typst_prefers_office_font_order_when_context_present() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Title".to_string(),
            style: TextStyle {
                font_family: Some("Pretendard".to_string()),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let context = FontSearchContext::for_test(
        Vec::new(),
        &["Apple SD Gothic Neo", "Malgun Gothic"],
        &["Malgun Gothic"],
        &[],
    );

    let output = generate_typst_with_options_and_font_context(
        &doc,
        &ConvertOptions::default(),
        Some(&context),
    )
    .unwrap();

    let apple_index = output
        .source
        .find("\"Apple SD Gothic Neo\"")
        .expect("Apple SD Gothic Neo should appear in Typst output");
    let malgun_index = output
        .source
        .find("\"Malgun Gothic\"")
        .expect("Malgun Gothic should appear in Typst output");
    assert!(
        malgun_index < apple_index,
        "Office-resolved font ordering should win in Typst output: {}",
        output.source
    );
}

#[test]
fn test_generate_heading_level_1() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            heading_level: Some(1),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Main Title".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#heading(level: 1)[Main Title]"),
        "H1 paragraph should emit #heading(level: 1): {result}"
    );
}

#[test]
fn test_generate_heading_level_2() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            heading_level: Some(2),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Sub Section".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#heading(level: 2)[Sub Section]"),
        "H2 paragraph should emit #heading(level: 2): {result}"
    );
}

#[test]
fn test_generate_heading_levels_3_to_6() {
    for level in 3..=6u8 {
        let text = format!("Heading {level}");
        let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            style: ParagraphStyle {
                heading_level: Some(level),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: text.clone(),
                style: TextStyle::default(),
                href: None,
                footnote: None,
            }],
        })])]);
        let result = generate_typst(&doc).unwrap().source;
        let expected = format!("#heading(level: {level})[{text}]");
        assert!(
            result.contains(&expected),
            "H{level} should emit {expected}: {result}"
        );
    }
}

#[test]
fn test_generate_heading_with_styled_run() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle {
            heading_level: Some(1),
            ..ParagraphStyle::default()
        },
        runs: vec![Run {
            text: "Styled Heading".to_string(),
            style: TextStyle {
                bold: Some(true),
                font_size: Some(24.0),
                ..TextStyle::default()
            },
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#heading(level: 1)"),
        "Heading with styling should still emit #heading: {result}"
    );
}

#[test]
fn test_generate_regular_paragraph_no_heading() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        style: ParagraphStyle::default(),
        runs: vec![Run {
            text: "Normal text".to_string(),
            style: TextStyle::default(),
            href: None,
            footnote: None,
        }],
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("#heading"),
        "Regular paragraph should not emit #heading: {result}"
    );
}
