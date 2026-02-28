//! PDF manipulation operations: merge, split, and page counting.
//!
//! These operations work on existing PDF files and are independent
//! from the document conversion pipeline.

use crate::error::ConvertError;
use lopdf::{Document, dictionary};

/// A range of pages to extract (1-indexed, inclusive).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PageRange {
    /// Start page (1-indexed, inclusive).
    pub start: u32,
    /// End page (1-indexed, inclusive).
    pub end: u32,
}

impl PageRange {
    /// Create a new page range (1-indexed, inclusive on both ends).
    pub fn new(start: u32, end: u32) -> Self {
        Self { start, end }
    }

    /// Parse a page range string like "1-5" or "3".
    pub fn parse(s: &str) -> Result<Self, String> {
        if let Some((start_str, end_str)) = s.split_once('-') {
            let start: u32 = start_str
                .trim()
                .parse()
                .map_err(|_| format!("invalid start page: {start_str}"))?;
            let end: u32 = end_str
                .trim()
                .parse()
                .map_err(|_| format!("invalid end page: {end_str}"))?;
            if start == 0 || end == 0 {
                return Err("page numbers must be >= 1".to_string());
            }
            if start > end {
                return Err(format!("start ({start}) must be <= end ({end})"));
            }
            Ok(Self::new(start, end))
        } else {
            let n: u32 = s
                .trim()
                .parse()
                .map_err(|_| format!("invalid page number: {s}"))?;
            if n == 0 {
                return Err("page number must be >= 1".to_string());
            }
            Ok(Self::new(n, n))
        }
    }
}

/// Count the number of pages in a PDF.
pub fn page_count(input: &[u8]) -> Result<u32, ConvertError> {
    let doc =
        Document::load_mem(input).map_err(|e| ConvertError::Parse(format!("invalid PDF: {e}")))?;
    Ok(doc.get_pages().len() as u32)
}

/// Merge multiple PDFs into a single PDF.
///
/// Each element of `inputs` is the raw bytes of a PDF file.
/// Returns the merged PDF bytes.
pub fn merge(inputs: &[&[u8]]) -> Result<Vec<u8>, ConvertError> {
    if inputs.is_empty() {
        return Err(ConvertError::Parse("no input PDFs to merge".to_string()));
    }

    if inputs.len() == 1 {
        // Single PDF — just return a copy
        return Ok(inputs[0].to_vec());
    }

    // Load all documents
    let documents: Vec<Document> = inputs
        .iter()
        .enumerate()
        .map(|(i, data)| {
            Document::load_mem(data)
                .map_err(|e| ConvertError::Parse(format!("invalid PDF at index {i}: {e}")))
        })
        .collect::<Result<_, _>>()?;

    // Use lopdf's merge approach: renumber objects, collect pages
    let mut max_id = 1;
    let mut all_pages = Vec::new();
    let mut all_objects = std::collections::BTreeMap::new();

    for mut doc in documents {
        doc.renumber_objects_with(max_id);
        max_id = doc.max_id + 1;

        // Collect page references in order
        let pages = doc.get_pages();
        let mut page_ids: Vec<_> = pages.into_iter().collect();
        page_ids.sort_by_key(|(num, _)| *num);
        for (_, page_id) in &page_ids {
            all_pages.push(*page_id);
        }

        // Collect all objects except Catalog
        for (id, object) in doc.objects {
            if let Ok(dict) = object.as_dict()
                && dict
                    .get(b"Type")
                    .ok()
                    .and_then(|t| t.as_name().ok())
                    .is_some_and(|name| name == b"Catalog")
            {
                continue;
            }
            all_objects.insert(id, object);
        }
    }

    // Build a new document with merged pages
    let mut merged = Document::with_version("1.7");

    // Insert all collected objects
    for (id, object) in &all_objects {
        merged.objects.insert(*id, object.clone());
    }
    merged.max_id = max_id;

    // Create Pages dictionary
    let pages_id = merged.new_object_id();
    let page_refs: Vec<lopdf::Object> = all_pages
        .iter()
        .map(|id| lopdf::Object::Reference(*id))
        .collect();

    let pages_dict = dictionary! {
        "Type" => "Pages",
        "Count" => all_pages.len() as i64,
        "Kids" => page_refs,
    };
    merged
        .objects
        .insert(pages_id, lopdf::Object::Dictionary(pages_dict));

    // Update each page's Parent reference
    for page_id in &all_pages {
        if let Ok(page_dict) = merged.objects.get_mut(page_id).unwrap().as_dict_mut() {
            page_dict.set("Parent", lopdf::Object::Reference(pages_id));
        }
    }

    // Create Catalog
    let catalog_id = merged.new_object_id();
    let catalog_dict = dictionary! {
        "Type" => "Catalog",
        "Pages" => pages_id,
    };
    merged
        .objects
        .insert(catalog_id, lopdf::Object::Dictionary(catalog_dict));
    merged
        .trailer
        .set("Root", lopdf::Object::Reference(catalog_id));

    // Remove orphaned intermediate Pages nodes from source documents
    // (they are no longer needed since we have a new top-level Pages node)
    let page_set: std::collections::HashSet<_> = all_pages.iter().collect();
    let mut to_remove = Vec::new();
    for (id, object) in &merged.objects {
        if page_set.contains(id) || *id == pages_id || *id == catalog_id {
            continue;
        }
        if let Ok(dict) = object.as_dict()
            && dict
                .get(b"Type")
                .ok()
                .and_then(|t| t.as_name().ok())
                .is_some_and(|name| name == b"Pages")
        {
            to_remove.push(*id);
        }
    }
    for id in to_remove {
        merged.objects.remove(&id);
    }

    merged.compress();

    let mut output = Vec::new();
    merged
        .save_to(&mut output)
        .map_err(|e| ConvertError::Render(format!("failed to write merged PDF: {e}")))?;

    Ok(output)
}

/// Split a PDF into multiple PDFs based on page ranges.
///
/// Each `PageRange` specifies a 1-indexed inclusive range of pages to extract.
/// Returns a vector of PDF byte arrays, one per range.
pub fn split(input: &[u8], ranges: &[PageRange]) -> Result<Vec<Vec<u8>>, ConvertError> {
    if ranges.is_empty() {
        return Err(ConvertError::Parse(
            "no page ranges specified for split".to_string(),
        ));
    }

    let doc =
        Document::load_mem(input).map_err(|e| ConvertError::Parse(format!("invalid PDF: {e}")))?;

    let total_pages = doc.get_pages().len() as u32;

    // Validate ranges
    for range in ranges {
        if range.start > total_pages || range.end > total_pages {
            return Err(ConvertError::Parse(format!(
                "page range {}-{} exceeds document page count ({total_pages})",
                range.start, range.end
            )));
        }
    }

    let mut results = Vec::with_capacity(ranges.len());

    for range in ranges {
        let mut split_doc = doc.clone();

        // Determine which pages to delete (all pages NOT in range)
        let pages_to_delete: Vec<u32> = (1..=total_pages)
            .filter(|p| *p < range.start || *p > range.end)
            .collect();

        if !pages_to_delete.is_empty() {
            split_doc.delete_pages(&pages_to_delete);
        }

        split_doc.compress();

        let mut output = Vec::new();
        split_doc
            .save_to(&mut output)
            .map_err(|e| ConvertError::Render(format!("failed to write split PDF: {e}")))?;

        results.push(output);
    }

    Ok(results)
}

#[cfg(test)]
mod tests {
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
}
