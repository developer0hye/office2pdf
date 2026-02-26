use std::collections::HashMap;
use std::io::{Cursor, Read};

use quick_xml::Reader;
use quick_xml::events::Event;
use zip::ZipArchive;

use crate::error::ConvertError;
use crate::ir::{
    Alignment, Block, Color, Document, FixedElement, FixedElementKind, FixedPage, Metadata, Page,
    PageSize, Paragraph, ParagraphStyle, Run, StyleSheet, TextStyle,
};
use crate::parser::Parser;

pub struct PptxParser;

/// Convert EMU (English Metric Units) to points.
/// 1 inch = 914400 EMU, 1 inch = 72 points, so 1 pt = 12700 EMU.
fn emu_to_pt(emu: i64) -> f64 {
    emu as f64 / 12700.0
}

impl Parser for PptxParser {
    fn parse(&self, data: &[u8]) -> Result<Document, ConvertError> {
        let cursor = Cursor::new(data);
        let mut archive = ZipArchive::new(cursor)
            .map_err(|e| ConvertError::Parse(format!("Failed to read PPTX: {e}")))?;

        // Read and parse presentation.xml for slide size and slide references
        let pres_xml = read_zip_entry(&mut archive, "ppt/presentation.xml")?;
        let (slide_size, slide_rids) = parse_presentation_xml(&pres_xml)?;

        // Read and parse presentation.xml.rels for rId → slide path mapping
        let rels_xml = read_zip_entry(&mut archive, "ppt/_rels/presentation.xml.rels")?;
        let rel_map = parse_rels_xml(&rels_xml);

        // Parse each slide in order
        let mut pages = Vec::new();
        for rid in &slide_rids {
            if let Some(target) = rel_map.get(rid) {
                let slide_path = if let Some(stripped) = target.strip_prefix('/') {
                    stripped.to_string()
                } else {
                    format!("ppt/{target}")
                };
                let slide_xml = read_zip_entry(&mut archive, &slide_path)?;
                let elements = parse_slide_xml(&slide_xml)?;
                pages.push(Page::Fixed(FixedPage {
                    size: slide_size,
                    elements,
                }));
            }
        }

        Ok(Document {
            metadata: Metadata::default(),
            pages,
            styles: StyleSheet::default(),
        })
    }
}

/// Read a file from the ZIP archive as a UTF-8 string.
fn read_zip_entry<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    path: &str,
) -> Result<String, ConvertError> {
    let mut file = archive
        .by_name(path)
        .map_err(|e| ConvertError::Parse(format!("Missing {path} in PPTX: {e}")))?;
    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|e| ConvertError::Parse(format!("Failed to read {path}: {e}")))?;
    Ok(content)
}

/// Parse presentation.xml to extract slide size and ordered slide relationship IDs.
fn parse_presentation_xml(xml: &str) -> Result<(PageSize, Vec<String>), ConvertError> {
    let mut reader = Reader::from_str(xml);
    // Default slide dimensions: 10" x 7.5" (standard 4:3)
    let mut slide_size = PageSize {
        width: 720.0,
        height: 540.0,
    };
    let mut slide_rids = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e)) => {
                handle_presentation_element(e, &mut slide_size, &mut slide_rids);
            }
            Ok(Event::Start(ref e)) => {
                handle_presentation_element(e, &mut slide_size, &mut slide_rids);
            }
            Ok(Event::Eof) => break,
            Err(e) => {
                return Err(ConvertError::Parse(format!(
                    "XML error in presentation.xml: {e}"
                )));
            }
            _ => {}
        }
    }

    Ok((slide_size, slide_rids))
}

/// Handle a single element from presentation.xml (both Start and Empty events).
fn handle_presentation_element(
    e: &quick_xml::events::BytesStart,
    slide_size: &mut PageSize,
    slide_rids: &mut Vec<String>,
) {
    match e.local_name().as_ref() {
        b"sldSz" => {
            let cx = get_attr_i64(e, b"cx").unwrap_or(9_144_000);
            let cy = get_attr_i64(e, b"cy").unwrap_or(6_858_000);
            *slide_size = PageSize {
                width: emu_to_pt(cx),
                height: emu_to_pt(cy),
            };
        }
        b"sldId" => {
            // r:id attribute contains the relationship ID
            if let Some(rid) = get_attr_str(e, b"r:id") {
                slide_rids.push(rid);
            }
        }
        _ => {}
    }
}

/// Parse a .rels file to build Id → Target mapping.
fn parse_rels_xml(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let mut reader = Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    if let (Some(id), Some(target)) =
                        (get_attr_str(e, b"Id"), get_attr_str(e, b"Target"))
                    {
                        map.insert(id, target);
                    }
                }
            }
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    if let (Some(id), Some(target)) =
                        (get_attr_str(e, b"Id"), get_attr_str(e, b"Target"))
                    {
                        map.insert(id, target);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    map
}

/// Parse a slide XML to extract positioned text box elements.
fn parse_slide_xml(xml: &str) -> Result<Vec<FixedElement>, ConvertError> {
    let mut reader = Reader::from_str(xml);
    let mut elements = Vec::new();

    // Shape-level state
    let mut in_shape = false;
    let mut shape_depth: usize = 0;
    let mut shape_x: i64 = 0;
    let mut shape_y: i64 = 0;
    let mut shape_cx: i64 = 0;
    let mut shape_cy: i64 = 0;

    // Transform state
    let mut in_xfrm = false;

    // Text body state
    let mut in_txbody = false;
    let mut has_txbody = false;
    let mut paragraphs: Vec<Paragraph> = Vec::new();

    // Paragraph state
    let mut in_para = false;
    let mut para_style = ParagraphStyle::default();
    let mut runs: Vec<Run> = Vec::new();

    // Run state
    let mut in_run = false;
    let mut run_style = TextStyle::default();
    let mut run_text = String::new();

    // Sub-element state
    let mut in_text = false;
    let mut in_rpr = false;
    let mut in_solid_fill = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"sp" if !in_shape => {
                        in_shape = true;
                        shape_depth = 1;
                        shape_x = 0;
                        shape_y = 0;
                        shape_cx = 0;
                        shape_cy = 0;
                        has_txbody = false;
                        paragraphs.clear();
                    }
                    b"sp" if in_shape => {
                        // Nested shape (inside group shape) — track depth
                        shape_depth += 1;
                    }
                    b"xfrm" if in_shape && !in_txbody => {
                        in_xfrm = true;
                    }
                    b"txBody" if in_shape => {
                        in_txbody = true;
                        has_txbody = true;
                    }
                    b"p" if in_txbody => {
                        in_para = true;
                        para_style = ParagraphStyle::default();
                        runs.clear();
                    }
                    b"pPr" if in_para && !in_run => {
                        // Paragraph properties (alignment, etc.)
                        extract_paragraph_props(e, &mut para_style);
                    }
                    b"r" if in_para => {
                        in_run = true;
                        run_style = TextStyle::default();
                        run_text.clear();
                    }
                    b"rPr" if in_run => {
                        in_rpr = true;
                        extract_rpr_attributes(e, &mut run_style);
                    }
                    b"solidFill" if in_rpr => {
                        in_solid_fill = true;
                    }
                    b"t" if in_run => {
                        in_text = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"off" if in_xfrm => {
                        shape_x = get_attr_i64(e, b"x").unwrap_or(0);
                        shape_y = get_attr_i64(e, b"y").unwrap_or(0);
                    }
                    b"ext" if in_xfrm => {
                        shape_cx = get_attr_i64(e, b"cx").unwrap_or(0);
                        shape_cy = get_attr_i64(e, b"cy").unwrap_or(0);
                    }
                    b"rPr" if in_run => {
                        extract_rpr_attributes(e, &mut run_style);
                    }
                    b"pPr" if in_para && !in_run => {
                        extract_paragraph_props(e, &mut para_style);
                    }
                    b"srgbClr" if in_solid_fill => {
                        if let Some(hex) = get_attr_str(e, b"val") {
                            run_style.color = parse_hex_color(&hex);
                        }
                    }
                    b"latin" if in_rpr => {
                        if let Some(typeface) = get_attr_str(e, b"typeface") {
                            run_style.font_family = Some(typeface);
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref t)) => {
                if in_text {
                    if let Ok(text) = t.xml_content() {
                        run_text.push_str(&text);
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"sp" if in_shape => {
                        shape_depth -= 1;
                        if shape_depth == 0 {
                            // Complete the shape — emit as TextBox if it had a text body
                            if has_txbody {
                                let blocks: Vec<Block> =
                                    paragraphs.drain(..).map(Block::Paragraph).collect();
                                elements.push(FixedElement {
                                    x: emu_to_pt(shape_x),
                                    y: emu_to_pt(shape_y),
                                    width: emu_to_pt(shape_cx),
                                    height: emu_to_pt(shape_cy),
                                    kind: FixedElementKind::TextBox(blocks),
                                });
                            }
                            in_shape = false;
                        }
                    }
                    b"xfrm" if in_xfrm => {
                        in_xfrm = false;
                    }
                    b"txBody" if in_txbody => {
                        in_txbody = false;
                    }
                    b"p" if in_para => {
                        paragraphs.push(Paragraph {
                            style: para_style.clone(),
                            runs: std::mem::take(&mut runs),
                        });
                        in_para = false;
                    }
                    b"r" if in_run => {
                        if !run_text.is_empty() {
                            runs.push(Run {
                                text: std::mem::take(&mut run_text),
                                style: run_style.clone(),
                            });
                        }
                        in_run = false;
                    }
                    b"rPr" if in_rpr => {
                        in_rpr = false;
                    }
                    b"solidFill" if in_solid_fill => {
                        in_solid_fill = false;
                    }
                    b"t" if in_text => {
                        in_text = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(ConvertError::Parse(format!("XML error in slide: {e}"))),
            _ => {}
        }
    }

    Ok(elements)
}

/// Extract paragraph alignment from `<a:pPr>` attributes.
fn extract_paragraph_props(e: &quick_xml::events::BytesStart, style: &mut ParagraphStyle) {
    if let Some(algn) = get_attr_str(e, b"algn") {
        style.alignment = match algn.as_str() {
            "l" => Some(Alignment::Left),
            "ctr" => Some(Alignment::Center),
            "r" => Some(Alignment::Right),
            "just" => Some(Alignment::Justify),
            _ => None,
        };
    }
}

/// Extract text formatting attributes from `<a:rPr>` element.
fn extract_rpr_attributes(e: &quick_xml::events::BytesStart, style: &mut TextStyle) {
    if let Some(val) = get_attr_str(e, b"b") {
        style.bold = Some(val == "1" || val == "true");
    }
    if let Some(val) = get_attr_str(e, b"i") {
        style.italic = Some(val == "1" || val == "true");
    }
    if let Some(val) = get_attr_str(e, b"u") {
        style.underline = Some(val != "none");
    }
    if let Some(val) = get_attr_str(e, b"strike") {
        style.strikethrough = Some(val != "noStrike");
    }
    if let Some(sz) = get_attr_i64(e, b"sz") {
        // Font size in hundredths of a point (e.g. 1200 = 12pt)
        style.font_size = Some(sz as f64 / 100.0);
    }
}

/// Get a string attribute value from an XML element.
/// Matches on full qualified name first (e.g. `r:id`), then local name.
fn get_attr_str(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<String> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == key || attr.key.local_name().as_ref() == key {
            return attr.unescape_value().ok().map(|v| v.to_string());
        }
    }
    None
}

/// Get an i64 attribute value from an XML element.
fn get_attr_i64(e: &quick_xml::events::BytesStart, key: &[u8]) -> Option<i64> {
    get_attr_str(e, key).and_then(|v| v.parse().ok())
}

/// Parse a 6-character hex color string (e.g. "FF0000") to an IR Color.
fn parse_hex_color(hex: &str) -> Option<Color> {
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some(Color::new(r, g, b))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Write;
    use zip::write::FileOptions;

    // ── Test helpers ─────────────────────────────────────────────────────

    /// Build a minimal PPTX file as bytes from slide XML strings.
    fn build_test_pptx(slide_cx_emu: i64, slide_cy_emu: i64, slide_xmls: &[String]) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = FileOptions::default();

        // [Content_Types].xml
        let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
        ct.push_str(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        );
        ct.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
        ct.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
        for i in 0..slide_xmls.len() {
            ct.push_str(&format!(
                r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
                i + 1
            ));
        }
        ct.push_str("</Types>");
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(ct.as_bytes()).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(
            br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        )
        .unwrap();

        // ppt/presentation.xml
        let mut pres = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{}" cy="{}"/><p:sldIdLst>"#,
            slide_cx_emu, slide_cy_emu
        );
        for i in 0..slide_xmls.len() {
            pres.push_str(&format!(
                r#"<p:sldId id="{}" r:id="rId{}"/>"#,
                256 + i,
                2 + i
            ));
        }
        pres.push_str("</p:sldIdLst></p:presentation>");
        zip.start_file("ppt/presentation.xml", opts).unwrap();
        zip.write_all(pres.as_bytes()).unwrap();

        // ppt/_rels/presentation.xml.rels
        let mut pres_rels = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for i in 0..slide_xmls.len() {
            pres_rels.push_str(&format!(
                r#"<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
                2 + i,
                1 + i
            ));
        }
        pres_rels.push_str("</Relationships>");
        zip.start_file("ppt/_rels/presentation.xml.rels", opts)
            .unwrap();
        zip.write_all(pres_rels.as_bytes()).unwrap();

        // Slides
        for (i, slide_xml) in slide_xmls.iter().enumerate() {
            zip.start_file(format!("ppt/slides/slide{}.xml", i + 1), opts)
                .unwrap();
            zip.write_all(slide_xml.as_bytes()).unwrap();
        }

        let cursor = zip.finish().unwrap();
        cursor.into_inner()
    }

    /// Create a slide XML with the given shape elements.
    fn make_slide_xml(shapes: &[String]) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
        );
        for shape in shapes {
            xml.push_str(shape);
        }
        xml.push_str("</p:spTree></p:cSld></p:sld>");
        xml
    }

    /// Create an empty slide XML (no shapes).
    fn make_empty_slide_xml() -> String {
        make_slide_xml(&[])
    }

    /// Create a simple text box shape XML.
    fn make_text_box(x: i64, y: i64, cx: i64, cy: i64, text: &str) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
        )
    }

    /// Create a text box with formatted text runs.
    fn make_formatted_text_box(x: i64, y: i64, cx: i64, cy: i64, runs_xml: &str) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p>{runs_xml}</a:p></p:txBody></p:sp>"#
        )
    }

    /// Create a text box with multiple paragraphs.
    fn make_multi_para_text_box(x: i64, y: i64, cx: i64, cy: i64, paragraphs_xml: &str) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/>{paragraphs_xml}</p:txBody></p:sp>"#
        )
    }

    /// Standard 4:3 slide size in EMU (10" x 7.5").
    const SLIDE_CX: i64 = 9_144_000;
    const SLIDE_CY: i64 = 6_858_000;

    /// Helper: get the first FixedPage from a Document.
    fn first_fixed_page(doc: &Document) -> &FixedPage {
        match &doc.pages[0] {
            Page::Fixed(p) => p,
            _ => panic!("Expected FixedPage"),
        }
    }

    /// Helper: get the TextBox blocks from a FixedElement.
    fn text_box_blocks(elem: &FixedElement) -> &[Block] {
        match &elem.kind {
            FixedElementKind::TextBox(blocks) => blocks,
            _ => panic!("Expected TextBox"),
        }
    }

    // ── Tests ────────────────────────────────────────────────────────────

    #[test]
    fn test_parse_empty_presentation() {
        // PPTX with zero slides → document with no pages
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();
        assert!(doc.pages.is_empty(), "Expected no pages");
    }

    #[test]
    fn test_parse_single_slide() {
        let slide = make_empty_slide_xml();
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();
        assert_eq!(doc.pages.len(), 1, "Expected 1 page");
        assert!(matches!(&doc.pages[0], Page::Fixed(_)));
    }

    #[test]
    fn test_slide_dimensions() {
        // 16:9 widescreen: 12192000 × 6858000 EMU = 960pt × 540pt
        let cx = 12_192_000i64;
        let cy = 6_858_000i64;
        let slide = make_empty_slide_xml();
        let data = build_test_pptx(cx, cy, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        let expected_w = cx as f64 / 12700.0;
        let expected_h = cy as f64 / 12700.0;
        assert!(
            (page.size.width - expected_w).abs() < 0.1,
            "Expected width ~{expected_w}pt, got {}",
            page.size.width
        );
        assert!(
            (page.size.height - expected_h).abs() < 0.1,
            "Expected height ~{expected_h}pt, got {}",
            page.size.height
        );
    }

    #[test]
    fn test_text_box_extraction() {
        let shape = make_text_box(0, 0, 1_000_000, 500_000, "Hello World");
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

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
        // Position: 1000000 EMU x, 500000 EMU y → ~78.74pt, ~39.37pt
        // Size: 5000000 EMU cx, 2000000 EMU cy → ~393.70pt, ~157.48pt
        let x = 1_000_000i64;
        let y = 500_000i64;
        let cx = 5_000_000i64;
        let cy = 2_000_000i64;
        let shape = make_text_box(x, y, cx, cy, "Positioned");
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        let elem = &page.elements[0];

        let expected_x = x as f64 / 12700.0;
        let expected_y = y as f64 / 12700.0;
        let expected_w = cx as f64 / 12700.0;
        let expected_h = cy as f64 / 12700.0;

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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        // sz="2400" means 24pt (hundredths of a point)
        let runs_xml = r#"<a:r><a:rPr sz="2400"/><a:t>Large text</a:t></a:r>"#;
        let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs[0].style.font_size, Some(24.0));
    }

    #[test]
    fn test_text_box_combined_formatting() {
        let runs_xml = r#"<a:r><a:rPr b="1" i="1" u="sng" strike="sngStrike" sz="1800"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>Styled text</a:t></a:r>"#;
        let shape = make_formatted_text_box(0, 0, 1_000_000, 500_000, runs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

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
    }

    #[test]
    fn test_multiple_text_boxes() {
        let shape1 = make_text_box(100_000, 100_000, 2_000_000, 500_000, "Box 1");
        let shape2 = make_text_box(100_000, 700_000, 2_000_000, 500_000, "Box 2");
        let slide = make_slide_xml(&[shape1, shape2]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 2, "Expected 2 text boxes");

        // Check content of each box
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
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

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
        let runs_xml = r#"<a:r><a:rPr b="1"/><a:t>Bold </a:t></a:r><a:r><a:rPr i="1"/><a:t>Italic</a:t></a:r>"#;
        let shape = make_formatted_text_box(0, 0, 2_000_000, 500_000, runs_xml);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

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
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        let blocks = text_box_blocks(&page.elements[0]);
        let para = match &blocks[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.style.alignment, Some(Alignment::Center));
    }

    #[test]
    fn test_parse_invalid_data() {
        let parser = PptxParser;
        let result = parser.parse(b"not a valid pptx file");
        assert!(result.is_err());
        match result.unwrap_err() {
            ConvertError::Parse(_) => {}
            other => panic!("Expected Parse error, got: {other:?}"),
        }
    }

    #[test]
    fn test_slide_default_dimensions_4x3() {
        // Standard 4:3: 9144000 × 6858000 EMU = 720pt × 540pt
        let slide = make_empty_slide_xml();
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        assert!(
            (page.size.width - 720.0).abs() < 0.1,
            "Expected width ~720pt, got {}",
            page.size.width
        );
        assert!(
            (page.size.height - 540.0).abs() < 0.1,
            "Expected height ~540pt, got {}",
            page.size.height
        );
    }
}
