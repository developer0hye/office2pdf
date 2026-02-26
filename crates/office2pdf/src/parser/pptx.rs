use std::collections::HashMap;
use std::io::{Cursor, Read};

use quick_xml::Reader;
use quick_xml::events::Event;
use zip::ZipArchive;

use crate::error::ConvertError;
use crate::ir::{
    Alignment, Block, BorderSide, Color, Document, FixedElement, FixedElementKind, FixedPage,
    ImageData, ImageFormat, Metadata, Page, PageSize, Paragraph, ParagraphStyle, Run, Shape,
    ShapeKind, StyleSheet, TextStyle,
};
use crate::parser::Parser;

/// Map from relationship ID → (image bytes, format).
type SlideImageMap = HashMap<String, (Vec<u8>, ImageFormat)>;

/// Context for which element a `<a:solidFill>` belongs to.
#[derive(Clone, Copy, PartialEq, Eq)]
enum SolidFillCtx {
    None,
    /// Fill color of the shape itself (inside `<p:spPr>`, not `<a:ln>`).
    ShapeFill,
    /// Stroke/border color (inside `<a:ln>`).
    LineFill,
    /// Text run color (inside `<a:rPr>`).
    RunFill,
}

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

                // Load images referenced by this slide
                let slide_images = load_slide_images(&slide_path, &mut archive);

                let elements = parse_slide_xml(&slide_xml, &slide_images)?;
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

/// Load images referenced by a slide from its .rels file and the ZIP archive.
///
/// Reads `<slide-dir>/_rels/<slide-filename>.rels`, finds image relationships,
/// and loads the corresponding image bytes from the ZIP.
fn load_slide_images<R: Read + std::io::Seek>(
    slide_path: &str,
    archive: &mut ZipArchive<R>,
) -> SlideImageMap {
    let mut images = SlideImageMap::new();

    // Build .rels path: ppt/slides/slide1.xml → ppt/slides/_rels/slide1.xml.rels
    let slide_rels_path = if let Some((dir, filename)) = slide_path.rsplit_once('/') {
        format!("{dir}/_rels/{filename}.rels")
    } else {
        format!("_rels/{slide_path}.rels")
    };

    let rels_xml = match read_zip_entry(archive, &slide_rels_path) {
        Ok(xml) => xml,
        Err(_) => return images, // No .rels file → no images
    };

    let slide_dir = slide_path.rsplit_once('/').map(|(d, _)| d).unwrap_or("");
    let rel_map = parse_rels_xml(&rels_xml);

    for (id, target) in &rel_map {
        let format = match image_format_from_ext(target) {
            Some(f) => f,
            None => continue, // Not an image relationship
        };

        // Resolve relative path (e.g., "../media/image1.png" → "ppt/media/image1.png")
        let image_path = if let Some(stripped) = target.strip_prefix('/') {
            stripped.to_string()
        } else {
            resolve_relative_path(slide_dir, target)
        };

        if let Ok(mut file) = archive.by_name(&image_path) {
            let mut data = Vec::new();
            if file.read_to_end(&mut data).is_ok() {
                images.insert(id.clone(), (data, format));
            }
        }
    }

    images
}

/// Resolve a relative path against a base directory.
/// e.g., base="ppt/slides", relative="../media/image1.png" → "ppt/media/image1.png"
fn resolve_relative_path(base_dir: &str, relative: &str) -> String {
    let mut parts: Vec<&str> = if base_dir.is_empty() {
        Vec::new()
    } else {
        base_dir.split('/').collect()
    };
    for component in relative.split('/') {
        match component {
            ".." => {
                parts.pop();
            }
            "." | "" => {}
            other => parts.push(other),
        }
    }
    parts.join("/")
}

/// Determine image format from file extension, or None if not a recognized image.
fn image_format_from_ext(path: &str) -> Option<ImageFormat> {
    let lower = path.to_lowercase();
    if lower.ends_with(".png") {
        Some(ImageFormat::Png)
    } else if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        Some(ImageFormat::Jpeg)
    } else if lower.ends_with(".gif") {
        Some(ImageFormat::Gif)
    } else if lower.ends_with(".bmp") {
        Some(ImageFormat::Bmp)
    } else if lower.ends_with(".tiff") || lower.ends_with(".tif") {
        Some(ImageFormat::Tiff)
    } else {
        None
    }
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
                if e.local_name().as_ref() == b"Relationship"
                    && let (Some(id), Some(target)) =
                        (get_attr_str(e, b"Id"), get_attr_str(e, b"Target"))
                {
                    map.insert(id, target);
                }
            }
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"Relationship"
                    && let (Some(id), Some(target)) =
                        (get_attr_str(e, b"Id"), get_attr_str(e, b"Target"))
                {
                    map.insert(id, target);
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    map
}

/// Parse a slide XML to extract positioned elements (text boxes, shapes, images).
fn parse_slide_xml(xml: &str, images: &SlideImageMap) -> Result<Vec<FixedElement>, ConvertError> {
    let mut reader = Reader::from_str(xml);
    let mut elements = Vec::new();

    // ── Shape-level state ────────────────────────────────────────────────
    let mut in_shape = false;
    let mut shape_depth: usize = 0;
    let mut shape_x: i64 = 0;
    let mut shape_y: i64 = 0;
    let mut shape_cx: i64 = 0;
    let mut shape_cy: i64 = 0;

    // Shape property state (geometry, fill, border)
    let mut in_sp_pr = false;
    let mut prst_geom: Option<String> = None;
    let mut shape_fill: Option<Color> = None;
    let mut in_ln = false;
    let mut ln_width_emu: i64 = 0;
    let mut ln_color: Option<Color> = None;

    // Transform state (for shapes)
    let mut in_xfrm = false;

    // Text body state
    let mut in_txbody = false;
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
    let mut solid_fill_ctx = SolidFillCtx::None;

    // ── Picture-level state ──────────────────────────────────────────────
    let mut in_pic = false;
    let mut pic_x: i64 = 0;
    let mut pic_y: i64 = 0;
    let mut pic_cx: i64 = 0;
    let mut pic_cy: i64 = 0;
    let mut blip_embed: Option<String> = None;
    let mut in_pic_xfrm = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    // ── Shape start ──────────────────────────────────
                    b"sp" if !in_shape && !in_pic => {
                        in_shape = true;
                        shape_depth = 1;
                        shape_x = 0;
                        shape_y = 0;
                        shape_cx = 0;
                        shape_cy = 0;
                        in_sp_pr = false;
                        prst_geom = None;
                        shape_fill = None;
                        in_ln = false;
                        ln_width_emu = 0;
                        ln_color = None;
                        in_txbody = false;
                        paragraphs.clear();
                    }
                    b"sp" if in_shape => {
                        shape_depth += 1;
                    }

                    // ── Shape properties ─────────────────────────────
                    b"spPr" if in_shape && !in_txbody => {
                        in_sp_pr = true;
                    }
                    b"xfrm" if in_shape && in_sp_pr => {
                        in_xfrm = true;
                    }
                    b"prstGeom" if in_sp_pr => {
                        if let Some(prst) = get_attr_str(e, b"prst") {
                            prst_geom = Some(prst);
                        }
                    }
                    b"solidFill" if in_sp_pr && !in_ln && !in_rpr => {
                        solid_fill_ctx = SolidFillCtx::ShapeFill;
                    }
                    b"ln" if in_sp_pr => {
                        in_ln = true;
                        ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                    }
                    b"solidFill" if in_ln => {
                        solid_fill_ctx = SolidFillCtx::LineFill;
                    }

                    // ── Text body ────────────────────────────────────
                    b"txBody" if in_shape => {
                        in_txbody = true;
                    }
                    b"p" if in_txbody => {
                        in_para = true;
                        para_style = ParagraphStyle::default();
                        runs.clear();
                    }
                    b"pPr" if in_para && !in_run => {
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
                        solid_fill_ctx = SolidFillCtx::RunFill;
                    }
                    b"t" if in_run => {
                        in_text = true;
                    }

                    // ── Picture start ────────────────────────────────
                    b"pic" if !in_shape && !in_pic => {
                        in_pic = true;
                        pic_x = 0;
                        pic_y = 0;
                        pic_cx = 0;
                        pic_cy = 0;
                        blip_embed = None;
                        in_pic_xfrm = false;
                    }
                    b"spPr" if in_pic => {
                        // Re-use nothing — just mark for xfrm detection below
                    }
                    b"xfrm" if in_pic => {
                        in_pic_xfrm = true;
                    }
                    b"blipFill" if in_pic => {}

                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    // ── Shape xfrm offset/extent ─────────────────────
                    b"off" if in_xfrm => {
                        shape_x = get_attr_i64(e, b"x").unwrap_or(0);
                        shape_y = get_attr_i64(e, b"y").unwrap_or(0);
                    }
                    b"ext" if in_xfrm => {
                        shape_cx = get_attr_i64(e, b"cx").unwrap_or(0);
                        shape_cy = get_attr_i64(e, b"cy").unwrap_or(0);
                    }

                    // ── Picture xfrm offset/extent ───────────────────
                    b"off" if in_pic_xfrm => {
                        pic_x = get_attr_i64(e, b"x").unwrap_or(0);
                        pic_y = get_attr_i64(e, b"y").unwrap_or(0);
                    }
                    b"ext" if in_pic_xfrm => {
                        pic_cx = get_attr_i64(e, b"cx").unwrap_or(0);
                        pic_cy = get_attr_i64(e, b"cy").unwrap_or(0);
                    }

                    // ── Blip (image reference) ───────────────────────
                    b"blip" if in_pic => {
                        blip_embed = get_attr_str(e, b"r:embed");
                    }

                    // ── Preset geometry (empty element) ──────────────
                    b"prstGeom" if in_sp_pr => {
                        if let Some(prst) = get_attr_str(e, b"prst") {
                            prst_geom = Some(prst);
                        }
                    }

                    // ── Line element (empty, no fill children) ───────
                    b"ln" if in_sp_pr => {
                        ln_width_emu = get_attr_i64(e, b"w").unwrap_or(12700);
                    }

                    // ── Color value ──────────────────────────────────
                    b"srgbClr" if solid_fill_ctx != SolidFillCtx::None => {
                        if let Some(hex) = get_attr_str(e, b"val") {
                            let color = parse_hex_color(&hex);
                            match solid_fill_ctx {
                                SolidFillCtx::ShapeFill => shape_fill = color,
                                SolidFillCtx::LineFill => ln_color = color,
                                SolidFillCtx::RunFill => run_style.color = color,
                                SolidFillCtx::None => {}
                            }
                        }
                    }

                    // ── Run properties (empty element) ───────────────
                    b"rPr" if in_run => {
                        extract_rpr_attributes(e, &mut run_style);
                    }
                    b"pPr" if in_para && !in_run => {
                        extract_paragraph_props(e, &mut para_style);
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
                if in_text && let Ok(text) = t.xml_content() {
                    run_text.push_str(&text);
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    // ── Shape end ────────────────────────────────────
                    b"sp" if in_shape => {
                        shape_depth -= 1;
                        if shape_depth == 0 {
                            let has_text = paragraphs.iter().any(|p| !p.runs.is_empty());

                            if has_text {
                                // TextBox — has visible text content
                                let blocks: Vec<Block> =
                                    paragraphs.drain(..).map(Block::Paragraph).collect();
                                elements.push(FixedElement {
                                    x: emu_to_pt(shape_x),
                                    y: emu_to_pt(shape_y),
                                    width: emu_to_pt(shape_cx),
                                    height: emu_to_pt(shape_cy),
                                    kind: FixedElementKind::TextBox(blocks),
                                });
                            } else if let Some(ref geom) = prst_geom {
                                // Shape — no text, but has geometry
                                let kind = prst_to_shape_kind(
                                    geom,
                                    emu_to_pt(shape_cx),
                                    emu_to_pt(shape_cy),
                                );
                                let stroke = ln_color.map(|color| BorderSide {
                                    width: ln_width_emu as f64 / 12700.0,
                                    color,
                                });
                                elements.push(FixedElement {
                                    x: emu_to_pt(shape_x),
                                    y: emu_to_pt(shape_y),
                                    width: emu_to_pt(shape_cx),
                                    height: emu_to_pt(shape_cy),
                                    kind: FixedElementKind::Shape(Shape {
                                        kind,
                                        fill: shape_fill,
                                        stroke,
                                    }),
                                });
                            }
                            in_shape = false;
                        }
                    }

                    // ── Shape sub-elements end ───────────────────────
                    b"spPr" if in_sp_pr => {
                        in_sp_pr = false;
                    }
                    b"xfrm" if in_xfrm => {
                        in_xfrm = false;
                    }
                    b"ln" if in_ln => {
                        in_ln = false;
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
                    b"solidFill" if solid_fill_ctx != SolidFillCtx::None => {
                        solid_fill_ctx = SolidFillCtx::None;
                    }
                    b"t" if in_text => {
                        in_text = false;
                    }

                    // ── Picture end ──────────────────────────────────
                    b"pic" if in_pic => {
                        if let Some(ref rid) = blip_embed
                            && let Some((data, format)) = images.get(rid)
                        {
                            elements.push(FixedElement {
                                x: emu_to_pt(pic_x),
                                y: emu_to_pt(pic_y),
                                width: emu_to_pt(pic_cx),
                                height: emu_to_pt(pic_cy),
                                kind: FixedElementKind::Image(ImageData {
                                    data: data.clone(),
                                    format: *format,
                                    width: Some(emu_to_pt(pic_cx)),
                                    height: Some(emu_to_pt(pic_cy)),
                                }),
                            });
                        }
                        in_pic = false;
                    }
                    b"xfrm" if in_pic_xfrm => {
                        in_pic_xfrm = false;
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

/// Map a PPTX preset geometry name to an IR ShapeKind.
fn prst_to_shape_kind(prst: &str, width: f64, height: f64) -> ShapeKind {
    match prst {
        "ellipse" => ShapeKind::Ellipse,
        "line" | "straightConnector1" => ShapeKind::Line {
            x2: width,
            y2: height,
        },
        // All rectangular-ish shapes → Rectangle
        _ => ShapeKind::Rectangle,
    }
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

    // ── Shape test helpers ───────────────────────────────────────────────

    /// Create a shape XML element with preset geometry, optional fill and border.
    #[allow(clippy::too_many_arguments)]
    fn make_shape(
        x: i64,
        y: i64,
        cx: i64,
        cy: i64,
        prst: &str,
        fill_hex: Option<&str>,
        border_width_emu: Option<i64>,
        border_hex: Option<&str>,
    ) -> String {
        let fill_xml = fill_hex
            .map(|h| format!(r#"<a:solidFill><a:srgbClr val="{h}"/></a:solidFill>"#))
            .unwrap_or_default();

        let ln_xml = match (border_width_emu, border_hex) {
            (Some(w), Some(h)) => {
                format!(r#"<a:ln w="{w}"><a:solidFill><a:srgbClr val="{h}"/></a:solidFill></a:ln>"#)
            }
            _ => String::new(),
        };

        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill_xml}{ln_xml}</p:spPr></p:sp>"#
        )
    }

    /// Helper: extract the Shape from a FixedElement or panic.
    fn get_shape(elem: &FixedElement) -> &Shape {
        match &elem.kind {
            FixedElementKind::Shape(s) => s,
            other => panic!("Expected Shape, got {other:?}"),
        }
    }

    // ── Image test helpers ───────────────────────────────────────────────

    /// Create a minimal valid BMP (1×1 pixel, red) for test images.
    fn make_test_bmp() -> Vec<u8> {
        let mut bmp = Vec::new();
        // BMP header (14 bytes)
        bmp.extend_from_slice(b"BM");
        bmp.extend_from_slice(&70u32.to_le_bytes()); // file size
        bmp.extend_from_slice(&0u32.to_le_bytes()); // reserved
        bmp.extend_from_slice(&54u32.to_le_bytes()); // pixel data offset
        // DIB header (40 bytes)
        bmp.extend_from_slice(&40u32.to_le_bytes()); // header size
        bmp.extend_from_slice(&1i32.to_le_bytes()); // width
        bmp.extend_from_slice(&1i32.to_le_bytes()); // height
        bmp.extend_from_slice(&1u16.to_le_bytes()); // planes
        bmp.extend_from_slice(&24u16.to_le_bytes()); // bpp
        bmp.extend_from_slice(&0u32.to_le_bytes()); // compression
        bmp.extend_from_slice(&16u32.to_le_bytes()); // image size
        bmp.extend_from_slice(&2835u32.to_le_bytes()); // h resolution
        bmp.extend_from_slice(&2835u32.to_le_bytes()); // v resolution
        bmp.extend_from_slice(&0u32.to_le_bytes()); // colors
        bmp.extend_from_slice(&0u32.to_le_bytes()); // important colors
        // Pixel data: 1 pixel (BGR) + 1 byte padding to align to 4 bytes
        bmp.extend_from_slice(&[0x00, 0x00, 0xFF, 0x00]);
        bmp
    }

    /// Create a picture XML element referencing an image via relationship ID.
    fn make_pic_xml(x: i64, y: i64, cx: i64, cy: i64, r_embed: &str) -> String {
        format!(
            r#"<p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{r_embed}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr></p:pic>"#
        )
    }

    /// Slide image for the test PPTX builder.
    struct TestSlideImage {
        rid: String,
        path: String,
        data: Vec<u8>,
    }

    /// Build a PPTX file with slides that have image relationships.
    fn build_test_pptx_with_images(
        slide_cx_emu: i64,
        slide_cy_emu: i64,
        slides: &[(String, Vec<TestSlideImage>)],
    ) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = FileOptions::default();

        // [Content_Types].xml
        let mut ct = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
        ct.push_str(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        );
        ct.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
        ct.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
        ct.push_str(r#"<Default Extension="png" ContentType="image/png"/>"#);
        ct.push_str(r#"<Default Extension="bmp" ContentType="image/bmp"/>"#);
        ct.push_str(r#"<Default Extension="jpeg" ContentType="image/jpeg"/>"#);
        for i in 0..slides.len() {
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
        for i in 0..slides.len() {
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
        for i in 0..slides.len() {
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

        // Slides and their .rels files
        for (i, (slide_xml, slide_images)) in slides.iter().enumerate() {
            let slide_num = i + 1;

            // Write slide XML
            zip.start_file(format!("ppt/slides/slide{slide_num}.xml"), opts)
                .unwrap();
            zip.write_all(slide_xml.as_bytes()).unwrap();

            // Write slide .rels if there are images
            if !slide_images.is_empty() {
                let mut rels = String::from(
                    r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                );
                for img in slide_images {
                    rels.push_str(&format!(
                        r#"<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{}"/>"#,
                        img.rid, img.path
                    ));
                }
                rels.push_str("</Relationships>");
                zip.start_file(format!("ppt/slides/_rels/slide{slide_num}.xml.rels"), opts)
                    .unwrap();
                zip.write_all(rels.as_bytes()).unwrap();

                // Write image media files
                for img in slide_images {
                    // Resolve the relative path (e.g., "../media/image1.png" → "ppt/media/image1.png")
                    let media_path = resolve_relative_path("ppt/slides", &img.path);
                    zip.start_file(media_path, opts).unwrap();
                    zip.write_all(&img.data).unwrap();
                }
            }
        }

        let cursor = zip.finish().unwrap();
        cursor.into_inner()
    }

    /// Helper: get the ImageData from a FixedElement or panic.
    fn get_image(elem: &FixedElement) -> &ImageData {
        match &elem.kind {
            FixedElementKind::Image(img) => img,
            other => panic!("Expected Image, got {other:?}"),
        }
    }

    // ── Shape tests ──────────────────────────────────────────────────────

    #[test]
    fn test_shape_rectangle_with_fill() {
        let shape = make_shape(
            1_000_000,
            500_000,
            3_000_000,
            2_000_000,
            "rect",
            Some("FF0000"),
            None,
            None,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1, "Expected 1 shape element");

        let elem = &page.elements[0];
        assert!((elem.x - emu_to_pt(1_000_000)).abs() < 0.1);
        assert!((elem.y - emu_to_pt(500_000)).abs() < 0.1);
        assert!((elem.width - emu_to_pt(3_000_000)).abs() < 0.1);
        assert!((elem.height - emu_to_pt(2_000_000)).abs() < 0.1);

        let shape = get_shape(elem);
        assert!(matches!(shape.kind, ShapeKind::Rectangle));
        assert_eq!(shape.fill, Some(Color::new(255, 0, 0)));
        assert!(shape.stroke.is_none());
    }

    #[test]
    fn test_shape_ellipse() {
        let shape = make_shape(
            0,
            0,
            2_000_000,
            2_000_000,
            "ellipse",
            Some("00FF00"),
            None,
            None,
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        assert!(matches!(s.kind, ShapeKind::Ellipse));
        assert_eq!(s.fill, Some(Color::new(0, 255, 0)));
    }

    #[test]
    fn test_shape_line() {
        let shape = make_shape(
            500_000,
            1_000_000,
            4_000_000,
            0,
            "line",
            None,
            Some(25400),
            Some("0000FF"),
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        match &s.kind {
            ShapeKind::Line { x2, y2 } => {
                assert!((*x2 - emu_to_pt(4_000_000)).abs() < 0.1);
                assert!((*y2 - 0.0).abs() < 0.1);
            }
            _ => panic!("Expected Line shape"),
        }
        assert!(s.fill.is_none());
        let stroke = s.stroke.as_ref().expect("Expected stroke on line");
        assert!((stroke.width - 2.0).abs() < 0.1); // 25400 EMU = 2pt
        assert_eq!(stroke.color, Color::new(0, 0, 255));
    }

    #[test]
    fn test_shape_with_fill_and_border() {
        let shape = make_shape(
            0,
            0,
            2_000_000,
            1_000_000,
            "rect",
            Some("FFFF00"),
            Some(12700),
            Some("000000"),
        );
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        assert_eq!(s.fill, Some(Color::new(255, 255, 0)));
        let stroke = s.stroke.as_ref().expect("Expected stroke");
        assert!((stroke.width - 1.0).abs() < 0.1); // 12700 EMU = 1pt
        assert_eq!(stroke.color, Color::black());
    }

    #[test]
    fn test_shape_no_fill_no_border() {
        let shape = make_shape(0, 0, 1_000_000, 1_000_000, "rect", None, None, None);
        let slide = make_slide_xml(&[shape]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        let s = get_shape(&page.elements[0]);
        assert!(s.fill.is_none());
        assert!(s.stroke.is_none());
    }

    #[test]
    fn test_multiple_shapes_on_slide() {
        let rect = make_shape(
            0,
            0,
            1_000_000,
            1_000_000,
            "rect",
            Some("FF0000"),
            None,
            None,
        );
        let ellipse = make_shape(
            2_000_000,
            0,
            1_000_000,
            1_000_000,
            "ellipse",
            Some("00FF00"),
            None,
            None,
        );
        let slide = make_slide_xml(&[rect, ellipse]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 2, "Expected 2 shape elements");
        assert!(matches!(
            get_shape(&page.elements[0]).kind,
            ShapeKind::Rectangle
        ));
        assert!(matches!(
            get_shape(&page.elements[1]).kind,
            ShapeKind::Ellipse
        ));
    }

    #[test]
    fn test_shapes_and_text_boxes_mixed() {
        let text_box = make_text_box(0, 0, 2_000_000, 500_000, "Hello");
        let rect = make_shape(
            0,
            1_000_000,
            2_000_000,
            500_000,
            "rect",
            Some("FF0000"),
            None,
            None,
        );
        let slide = make_slide_xml(&[text_box, rect]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 2, "Expected 2 elements");
        assert!(matches!(
            &page.elements[0].kind,
            FixedElementKind::TextBox(_)
        ));
        assert!(matches!(&page.elements[1].kind, FixedElementKind::Shape(_)));
    }

    // ── Image tests ──────────────────────────────────────────────────────

    #[test]
    fn test_image_basic_extraction() {
        let bmp_data = make_test_bmp();
        let pic = make_pic_xml(1_000_000, 500_000, 3_000_000, 2_000_000, "rId3");
        let slide_xml = make_slide_xml(&[pic]);
        let slide_images = vec![TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data.clone(),
        }];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 1, "Expected 1 image element");

        let elem = &page.elements[0];
        assert!((elem.x - emu_to_pt(1_000_000)).abs() < 0.1);
        assert!((elem.y - emu_to_pt(500_000)).abs() < 0.1);
        assert!((elem.width - emu_to_pt(3_000_000)).abs() < 0.1);
        assert!((elem.height - emu_to_pt(2_000_000)).abs() < 0.1);

        let img = get_image(elem);
        assert!(!img.data.is_empty(), "Image data should not be empty");
        assert_eq!(img.data, bmp_data);
    }

    #[test]
    fn test_image_format_detection() {
        let bmp_data = make_test_bmp();

        // Test BMP format
        let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
        let slide_xml = make_slide_xml(&[pic]);
        let slide_images = vec![TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data,
        }];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        let img = get_image(&page.elements[0]);
        assert_eq!(img.format, ImageFormat::Bmp);
    }

    #[test]
    fn test_image_dimensions_preserved() {
        let bmp_data = make_test_bmp();
        // 200pt × 100pt → 200*12700=2540000, 100*12700=1270000 EMU
        let pic = make_pic_xml(0, 0, 2_540_000, 1_270_000, "rId3");
        let slide_xml = make_slide_xml(&[pic]);
        let slide_images = vec![TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data,
        }];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        let img = get_image(&page.elements[0]);
        let w = img.width.expect("Expected width");
        let h = img.height.expect("Expected height");
        assert!((w - 200.0).abs() < 0.1, "Expected ~200pt, got {w}");
        assert!((h - 100.0).abs() < 0.1, "Expected ~100pt, got {h}");
    }

    #[test]
    fn test_image_with_shapes_and_text() {
        let bmp_data = make_test_bmp();
        let text_box = make_text_box(0, 0, 2_000_000, 500_000, "Title");
        let rect = make_shape(
            0,
            600_000,
            1_000_000,
            500_000,
            "rect",
            Some("AABBCC"),
            None,
            None,
        );
        let pic = make_pic_xml(2_000_000, 600_000, 1_500_000, 1_000_000, "rId3");
        let slide_xml = make_slide_xml(&[text_box, rect, pic]);
        let slide_images = vec![TestSlideImage {
            rid: "rId3".to_string(),
            path: "../media/image1.bmp".to_string(),
            data: bmp_data,
        }];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 3, "Expected 3 elements");
        assert!(matches!(
            &page.elements[0].kind,
            FixedElementKind::TextBox(_)
        ));
        assert!(matches!(&page.elements[1].kind, FixedElementKind::Shape(_)));
        assert!(matches!(&page.elements[2].kind, FixedElementKind::Image(_)));
    }

    #[test]
    fn test_image_missing_rid_ignored() {
        // Picture references rId3 but no image data for that rId
        let pic = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId99");
        let slide_xml = make_slide_xml(&[pic]);
        let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide_xml]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(
            page.elements.len(),
            0,
            "Missing image ref should be skipped"
        );
    }

    #[test]
    fn test_multiple_images_on_slide() {
        let bmp_data = make_test_bmp();
        let pic1 = make_pic_xml(0, 0, 1_000_000, 1_000_000, "rId3");
        let pic2 = make_pic_xml(2_000_000, 0, 1_500_000, 1_000_000, "rId4");
        let slide_xml = make_slide_xml(&[pic1, pic2]);
        let slide_images = vec![
            TestSlideImage {
                rid: "rId3".to_string(),
                path: "../media/image1.bmp".to_string(),
                data: bmp_data.clone(),
            },
            TestSlideImage {
                rid: "rId4".to_string(),
                path: "../media/image2.bmp".to_string(),
                data: bmp_data,
            },
        ];
        let data = build_test_pptx_with_images(SLIDE_CX, SLIDE_CY, &[(slide_xml, slide_images)]);
        let parser = PptxParser;
        let doc = parser.parse(&data).unwrap();

        let page = first_fixed_page(&doc);
        assert_eq!(page.elements.len(), 2, "Expected 2 image elements");
        assert!(matches!(&page.elements[0].kind, FixedElementKind::Image(_)));
        assert!(matches!(&page.elements[1].kind, FixedElementKind::Image(_)));
    }
}
