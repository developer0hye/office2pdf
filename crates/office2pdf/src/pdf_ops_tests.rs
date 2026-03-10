use super::*;

/// Create a minimal valid PDF with the given number of pages.
/// Each page is a simple blank A4 page.
fn make_test_pdf(num_pages: u32) -> Vec<u8> {
    let mut doc = Document::with_version("1.7");

    let pages_id = doc.new_object_id();
    let mut page_ids = Vec::new();

    for i in 0..num_pages {
        // Create a content stream with a simple text marker
        let content = format!("BT /F1 12 Tf 100 700 Td (Page {}) Tj ET", i + 1);
        let content_id =
            doc.add_object(lopdf::Stream::new(dictionary! {}, content.into_bytes()));

        let page_id = doc.add_object(dictionary! {
            "Type" => "Page",
            "Parent" => pages_id,
            "MediaBox" => vec![0.into(), 0.into(), 595.into(), 842.into()],
            "Contents" => content_id,
        });
        page_ids.push(page_id);
    }

    let page_refs: Vec<lopdf::Object> = page_ids
        .iter()
        .map(|id| lopdf::Object::Reference(*id))
        .collect();

    doc.objects.insert(
        pages_id,
        lopdf::Object::Dictionary(dictionary! {
            "Type" => "Pages",
            "Count" => num_pages as i64,
            "Kids" => page_refs,
        }),
    );

    let catalog_id = doc.add_object(dictionary! {
        "Type" => "Catalog",
        "Pages" => pages_id,
    });
    doc.trailer
        .set("Root", lopdf::Object::Reference(catalog_id));

    let mut output = Vec::new();
    doc.save_to(&mut output).unwrap();
    output
}

// --- PageRange tests ---

#[test]
fn test_page_range_new() {
    let r = PageRange::new(1, 5);
    assert_eq!(r.start, 1);
    assert_eq!(r.end, 5);
}

#[test]
fn test_page_range_parse_range() {
    let r = PageRange::parse("2-5").unwrap();
    assert_eq!(r.start, 2);
    assert_eq!(r.end, 5);
}

#[test]
fn test_page_range_parse_single() {
    let r = PageRange::parse("3").unwrap();
    assert_eq!(r.start, 3);
    assert_eq!(r.end, 3);
}

#[test]
fn test_page_range_parse_errors() {
    assert!(PageRange::parse("abc").is_err());
    assert!(PageRange::parse("0").is_err());
    assert!(PageRange::parse("5-2").is_err());
    assert!(PageRange::parse("0-3").is_err());
}

// --- page_count tests ---

#[test]
fn test_page_count_single_page() {
    let pdf = make_test_pdf(1);
    assert_eq!(page_count(&pdf).unwrap(), 1);
}

#[test]
fn test_page_count_multi_page() {
    let pdf = make_test_pdf(4);
    assert_eq!(page_count(&pdf).unwrap(), 4);
}

#[test]
fn test_page_count_invalid_pdf() {
    let result = page_count(b"not a pdf");
    assert!(result.is_err());
}

// --- merge tests ---

#[test]
fn test_merge_two_single_page_pdfs() {
    let pdf1 = make_test_pdf(1);
    let pdf2 = make_test_pdf(1);
    let merged = merge(&[&pdf1, &pdf2]).unwrap();

    // Merged PDF should have 2 pages
    assert_eq!(page_count(&merged).unwrap(), 2);
}

#[test]
fn test_merge_different_page_counts() {
    let pdf1 = make_test_pdf(2);
    let pdf2 = make_test_pdf(3);
    let merged = merge(&[&pdf1, &pdf2]).unwrap();

    assert_eq!(page_count(&merged).unwrap(), 5);
}

#[test]
fn test_merge_single_pdf_returns_copy() {
    let pdf = make_test_pdf(3);
    let merged = merge(&[&pdf]).unwrap();

    assert_eq!(page_count(&merged).unwrap(), 3);
}

#[test]
fn test_merge_three_pdfs() {
    let pdf1 = make_test_pdf(1);
    let pdf2 = make_test_pdf(2);
    let pdf3 = make_test_pdf(1);
    let merged = merge(&[&pdf1, &pdf2, &pdf3]).unwrap();

    assert_eq!(page_count(&merged).unwrap(), 4);
}

#[test]
fn test_merge_empty_input() {
    let result = merge(&[]);
    assert!(result.is_err());
}

#[test]
fn test_merge_invalid_pdf() {
    let valid = make_test_pdf(1);
    let result = merge(&[b"not a pdf" as &[u8], &valid]);
    assert!(result.is_err());
}

#[test]
fn test_merge_result_is_valid_pdf() {
    let pdf1 = make_test_pdf(1);
    let pdf2 = make_test_pdf(1);
    let merged = merge(&[&pdf1, &pdf2]).unwrap();

    // Should be loadable as a valid PDF
    let doc = Document::load_mem(&merged).unwrap();
    assert_eq!(doc.get_pages().len(), 2);
}

// --- split tests ---

#[test]
fn test_split_into_halves() {
    let pdf = make_test_pdf(4);
    let ranges = vec![PageRange::new(1, 2), PageRange::new(3, 4)];
    let parts = split(&pdf, &ranges).unwrap();

    assert_eq!(parts.len(), 2);
    assert_eq!(page_count(&parts[0]).unwrap(), 2);
    assert_eq!(page_count(&parts[1]).unwrap(), 2);
}

#[test]
fn test_split_single_page() {
    let pdf = make_test_pdf(3);
    let ranges = vec![PageRange::new(2, 2)];
    let parts = split(&pdf, &ranges).unwrap();

    assert_eq!(parts.len(), 1);
    assert_eq!(page_count(&parts[0]).unwrap(), 1);
}

#[test]
fn test_split_all_pages_individually() {
    let pdf = make_test_pdf(3);
    let ranges = vec![
        PageRange::new(1, 1),
        PageRange::new(2, 2),
        PageRange::new(3, 3),
    ];
    let parts = split(&pdf, &ranges).unwrap();

    assert_eq!(parts.len(), 3);
    for part in &parts {
        assert_eq!(page_count(part).unwrap(), 1);
    }
}

#[test]
fn test_split_empty_ranges() {
    let pdf = make_test_pdf(2);
    let result = split(&pdf, &[]);
    assert!(result.is_err());
}

#[test]
fn test_split_range_exceeds_page_count() {
    let pdf = make_test_pdf(2);
    let ranges = vec![PageRange::new(1, 5)];
    let result = split(&pdf, &ranges);
    assert!(result.is_err());
}

#[test]
fn test_split_invalid_pdf() {
    let result = split(b"not a pdf", &[PageRange::new(1, 1)]);
    assert!(result.is_err());
}

#[test]
fn test_split_results_are_valid_pdfs() {
    let pdf = make_test_pdf(4);
    let ranges = vec![PageRange::new(1, 2), PageRange::new(3, 4)];
    let parts = split(&pdf, &ranges).unwrap();

    for part in &parts {
        let doc = Document::load_mem(part).unwrap();
        assert_eq!(doc.get_pages().len(), 2);
    }
}

// --- Round-trip test: split then merge ---

#[test]
fn test_split_and_merge_round_trip() {
    let original = make_test_pdf(4);
    let ranges = vec![PageRange::new(1, 2), PageRange::new(3, 4)];
    let parts = split(&original, &ranges).unwrap();

    let merged = merge(&[&parts[0], &parts[1]]).unwrap();
    assert_eq!(page_count(&merged).unwrap(), 4);
}
