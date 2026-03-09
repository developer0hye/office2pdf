use super::*;

#[test]
fn test_generate_typst_with_metadata_title_and_author() {
    let doc = Document {
        metadata: Metadata {
            title: Some("Test Title".to_string()),
            author: Some("Test Author".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#set document(title: \"Test Title\", author: \"Test Author\")"),
        "Expected document metadata in Typst output, got: {result}"
    );
}

#[test]
fn test_generate_typst_with_metadata_title_only() {
    let doc = Document {
        metadata: Metadata {
            title: Some("Only Title".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("#set document(title: \"Only Title\")"),
        "Expected title-only metadata in Typst output, got: {result}"
    );
}

#[test]
fn test_generate_typst_without_metadata() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Paragraph(Paragraph {
        runs: vec![Run {
            text: "Hello".to_string(),
            style: TextStyle::default(),
            footnote: None,
            href: None,
        }],
        style: ParagraphStyle::default(),
    })])]);
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("#set document("),
        "Should not emit #set document when no metadata, got: {result}"
    );
}

#[test]
fn test_generate_typst_with_metadata_created_date() {
    let doc = Document {
        metadata: Metadata {
            title: Some("Dated Doc".to_string()),
            created: Some("2024-06-15T10:30:00Z".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("date: datetime(year: 2024, month: 6, day: 15"),
        "Expected document date from metadata created field, got: {result}"
    );
}

#[test]
fn test_generate_typst_with_metadata_date_only() {
    let doc = Document {
        metadata: Metadata {
            created: Some("2023-12-25T08:00:00Z".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        result.contains("date: datetime(year: 2023, month: 12, day: 25"),
        "Expected document date even without title/author, got: {result}"
    );
}

#[test]
fn test_generate_typst_with_invalid_created_date() {
    let doc = Document {
        metadata: Metadata {
            title: Some("Bad Date Doc".to_string()),
            created: Some("not-a-date".to_string()),
            ..Default::default()
        },
        pages: vec![make_flow_page(vec![Block::Paragraph(Paragraph {
            runs: vec![Run {
                text: "Hello".to_string(),
                style: TextStyle::default(),
                footnote: None,
                href: None,
            }],
            style: ParagraphStyle::default(),
        })])],
        styles: StyleSheet::default(),
    };
    let result = generate_typst(&doc).unwrap().source;
    assert!(
        !result.contains("date: datetime("),
        "Invalid date should not produce document date, got: {result}"
    );
}

#[test]
fn test_parse_iso8601_date_full() {
    let result = parse_iso8601_date("2024-06-15T10:30:45Z");
    assert_eq!(result, Some((2024, 6, 15, 10, 30, 45)));
}

#[test]
fn test_parse_iso8601_date_date_only() {
    let result = parse_iso8601_date("2023-12-25");
    assert_eq!(result, Some((2023, 12, 25, 0, 0, 0)));
}

#[test]
fn test_parse_iso8601_date_invalid() {
    assert_eq!(parse_iso8601_date("not-a-date"), None);
    assert_eq!(parse_iso8601_date(""), None);
    assert_eq!(parse_iso8601_date("2024"), None);
    assert_eq!(parse_iso8601_date("2024-13-01T00:00:00Z"), None);
    assert_eq!(parse_iso8601_date("2024-00-01T00:00:00Z"), None);
}
