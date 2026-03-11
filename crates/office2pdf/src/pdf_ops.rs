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

/// Load a PDF document from raw bytes, mapping errors to `ConvertError`.
fn load_pdf_document(input: &[u8], context: &str) -> Result<Document, ConvertError> {
    Document::load_mem(input).map_err(|e| ConvertError::Parse(format!("invalid PDF{context}: {e}")))
}

/// Validate that all page ranges fall within the document's page count.
fn validate_page_ranges(ranges: &[PageRange], total_pages: u32) -> Result<(), ConvertError> {
    for range in ranges {
        if range.start > total_pages || range.end > total_pages {
            return Err(ConvertError::Parse(format!(
                "page range {}-{} exceeds document page count ({total_pages})",
                range.start, range.end
            )));
        }
    }
    Ok(())
}

/// Compress a document and serialize it to PDF bytes.
fn save_pdf_to_bytes(doc: &mut Document, context: &str) -> Result<Vec<u8>, ConvertError> {
    doc.compress();
    let mut output: Vec<u8> = Vec::new();
    doc.save_to(&mut output)
        .map_err(|e| ConvertError::Render(format!("failed to write {context} PDF: {e}")))?;
    Ok(output)
}

/// Count the number of pages in a PDF.
pub fn page_count(input: &[u8]) -> Result<u32, ConvertError> {
    let doc: Document = load_pdf_document(input, "")?;
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
        .map(|(i, data)| load_pdf_document(data, &format!(" at index {i}")))
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
        if let Some(object) = merged.objects.get_mut(page_id)
            && let Ok(page_dict) = object.as_dict_mut()
        {
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

    save_pdf_to_bytes(&mut merged, "merged")
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

    let doc: Document = load_pdf_document(input, "")?;

    let total_pages: u32 = doc.get_pages().len() as u32;
    validate_page_ranges(ranges, total_pages)?;

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

        results.push(save_pdf_to_bytes(&mut split_doc, "split")?);
    }

    Ok(results)
}

#[cfg(test)]
#[path = "pdf_ops_tests.rs"]
mod tests;
