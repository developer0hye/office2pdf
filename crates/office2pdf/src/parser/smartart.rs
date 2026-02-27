use quick_xml::Reader;
/// SmartArt diagram parser for PPTX files.
///
/// Parses SmartArt data model XML to extract text content from diagram nodes.
/// SmartArt diagrams use the DrawingML Diagram namespace
/// (`http://schemas.openxmlformats.org/drawingml/2006/diagram`).
use quick_xml::events::Event;

/// Parse SmartArt data model XML and extract text items from data points.
///
/// The data model XML contains `<dgm:pt>` elements with `type` attributes.
/// We extract text from `type="node"` points (the actual content nodes),
/// skipping `type="doc"` (root), `type="pres"` (presentation), and
/// `type="parTrans"`/`type="sibTrans"` (transition) points.
///
/// Text is found inside `<dgm:t>/<a:p>/<a:r>/<a:t>` within each point.
pub(crate) fn parse_smartart_data_xml(xml: &str) -> Vec<String> {
    let mut items = Vec::new();
    let mut reader = Reader::from_str(xml);

    // State tracking
    let mut in_pt = false;
    let mut pt_is_node = false;
    let mut in_t_block = false; // inside <dgm:t> (text body)
    let mut in_a_r = false; // inside <a:r> (run)
    let mut in_a_t = false; // inside <a:t> (text)
    let mut current_text = String::new();
    let mut pt_depth: u32 = 0;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"pt" if !in_pt => {
                        in_pt = true;
                        pt_depth = 1;
                        current_text.clear();

                        // Check type attribute — default to "node" if absent
                        let mut pt_type = String::from("node");
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"type"
                                && let Ok(v) = attr.unescape_value()
                            {
                                pt_type = v.to_string();
                            }
                        }
                        pt_is_node = pt_type == "node";
                    }
                    b"pt" if in_pt => {
                        pt_depth += 1;
                    }
                    b"t" if in_a_r => {
                        // <a:t> inside <a:r> — the actual text element
                        in_a_t = true;
                    }
                    b"r" if in_t_block => {
                        in_a_r = true;
                    }
                    b"t" if in_pt && pt_is_node && !in_t_block => {
                        // <dgm:t> — the text body container
                        in_t_block = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Text(ref t)) if in_a_t => {
                if let Ok(text) = t.xml_content() {
                    current_text.push_str(&text);
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"pt" if in_pt => {
                        pt_depth -= 1;
                        if pt_depth == 0 {
                            if pt_is_node {
                                let trimmed = current_text.trim().to_string();
                                if !trimmed.is_empty() {
                                    items.push(trimmed);
                                }
                            }
                            in_pt = false;
                            current_text.clear();
                        }
                    }
                    b"t" if in_a_r => {
                        in_a_t = false;
                    }
                    b"r" if in_t_block => {
                        in_a_r = false;
                    }
                    b"t" if in_pt && !in_a_r => {
                        in_t_block = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    items
}

/// Reference to a SmartArt found in a slide's graphicFrame.
///
/// Contains the position/size of the graphicFrame and the relationship
/// ID for the SmartArt data model file.
#[derive(Debug)]
pub(crate) struct SmartArtRef {
    /// X position in EMU.
    pub x: i64,
    /// Y position in EMU.
    pub y: i64,
    /// Width in EMU.
    pub cx: i64,
    /// Height in EMU.
    pub cy: i64,
    /// Relationship ID for the data model (r:dm from dgm:relIds).
    pub data_rid: String,
}

/// Scan slide XML for SmartArt references within graphicFrame elements.
///
/// Returns a list of SmartArt references with position info and
/// the relationship ID needed to resolve the data file from .rels.
pub(crate) fn scan_smartart_refs(slide_xml: &str) -> Vec<SmartArtRef> {
    let mut refs = Vec::new();
    let mut reader = Reader::from_str(slide_xml);

    let mut in_graphic_frame = false;
    let mut gf_x: i64 = 0;
    let mut gf_y: i64 = 0;
    let mut gf_cx: i64 = 0;
    let mut gf_cy: i64 = 0;
    let mut in_gf_xfrm = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"graphicFrame" if !in_graphic_frame => {
                        in_graphic_frame = true;
                        gf_x = 0;
                        gf_y = 0;
                        gf_cx = 0;
                        gf_cy = 0;
                        in_gf_xfrm = false;
                    }
                    b"xfrm" if in_graphic_frame => {
                        in_gf_xfrm = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"off" if in_gf_xfrm => {
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"x" => {
                                    if let Ok(v) = attr.unescape_value() {
                                        gf_x = v.parse().unwrap_or(0);
                                    }
                                }
                                b"y" => {
                                    if let Ok(v) = attr.unescape_value() {
                                        gf_y = v.parse().unwrap_or(0);
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                    b"ext" if in_gf_xfrm => {
                        for attr in e.attributes().flatten() {
                            match attr.key.local_name().as_ref() {
                                b"cx" => {
                                    if let Ok(v) = attr.unescape_value() {
                                        gf_cx = v.parse().unwrap_or(0);
                                    }
                                }
                                b"cy" => {
                                    if let Ok(v) = attr.unescape_value() {
                                        gf_cy = v.parse().unwrap_or(0);
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                    b"relIds" if in_graphic_frame => {
                        // <dgm:relIds r:dm="rIdN" .../>
                        let mut data_rid = None;
                        for attr in e.attributes().flatten() {
                            // r:dm is the data model relationship
                            if attr.key.as_ref() == b"r:dm"
                                && let Ok(v) = attr.unescape_value()
                            {
                                data_rid = Some(v.to_string());
                            }
                        }
                        if let Some(rid) = data_rid {
                            refs.push(SmartArtRef {
                                x: gf_x,
                                y: gf_y,
                                cx: gf_cx,
                                cy: gf_cy,
                                data_rid: rid,
                            });
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"graphicFrame" if in_graphic_frame => {
                        in_graphic_frame = false;
                    }
                    b"xfrm" if in_gf_xfrm => {
                        in_gf_xfrm = false;
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    refs
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_smartart_data_basic_nodes() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dgm:ptLst>
            <dgm:pt modelId="0" type="doc">
              <dgm:prSet/>
              <dgm:spPr/>
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Root</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="1" type="node">
              <dgm:prSet/>
              <dgm:spPr/>
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Step 1</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="2" type="node">
              <dgm:prSet/>
              <dgm:spPr/>
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Step 2</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="3" type="node">
              <dgm:prSet/>
              <dgm:spPr/>
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Step 3</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
          </dgm:ptLst>
        </dgm:dataModel>"#;

        let items = parse_smartart_data_xml(xml);
        assert_eq!(items, vec!["Step 1", "Step 2", "Step 3"]);
    }

    #[test]
    fn test_parse_smartart_data_skips_transitions() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dgm:ptLst>
            <dgm:pt modelId="0" type="doc">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Root</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="1" type="node">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Item A</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="10" type="parTrans">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Trans</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="11" type="sibTrans">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>SibTrans</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="2" type="node">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Item B</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
          </dgm:ptLst>
        </dgm:dataModel>"#;

        let items = parse_smartart_data_xml(xml);
        assert_eq!(items, vec!["Item A", "Item B"]);
    }

    #[test]
    fn test_parse_smartart_data_empty_text_skipped() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dgm:ptLst>
            <dgm:pt modelId="1" type="node">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>  </a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="2" type="node">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Valid</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
          </dgm:ptLst>
        </dgm:dataModel>"#;

        let items = parse_smartart_data_xml(xml);
        assert_eq!(items, vec!["Valid"]);
    }

    #[test]
    fn test_parse_smartart_data_multi_run_text() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dgm:ptLst>
            <dgm:pt modelId="1" type="node">
              <dgm:t><a:bodyPr/><a:p>
                <a:r><a:t>Hello </a:t></a:r>
                <a:r><a:t>World</a:t></a:r>
              </a:p></dgm:t>
            </dgm:pt>
          </dgm:ptLst>
        </dgm:dataModel>"#;

        let items = parse_smartart_data_xml(xml);
        assert_eq!(items, vec!["Hello World"]);
    }

    #[test]
    fn test_parse_smartart_data_node_without_type_defaults_to_node() {
        // Points without an explicit type attribute default to "node"
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dgm:ptLst>
            <dgm:pt modelId="1">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Implicit Node</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
          </dgm:ptLst>
        </dgm:dataModel>"#;

        let items = parse_smartart_data_xml(xml);
        assert_eq!(items, vec!["Implicit Node"]);
    }

    #[test]
    fn test_scan_smartart_refs_basic() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <p:cSld><p:spTree>
            <p:graphicFrame>
              <p:nvGraphicFramePr>
                <p:cNvPr id="4" name="SmartArt"/>
                <p:cNvGraphicFramePr/>
                <p:nvPr/>
              </p:nvGraphicFramePr>
              <p:xfrm>
                <a:off x="914400" y="1828800"/>
                <a:ext cx="5486400" cy="3086100"/>
              </p:xfrm>
              <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                  <dgm:relIds r:dm="rId5" r:lo="rId6" r:qs="rId7" r:cs="rId8"/>
                </a:graphicData>
              </a:graphic>
            </p:graphicFrame>
          </p:spTree></p:cSld>
        </p:sld>"#;

        let refs = scan_smartart_refs(slide_xml);
        assert_eq!(refs.len(), 1);
        assert_eq!(refs[0].x, 914400);
        assert_eq!(refs[0].y, 1828800);
        assert_eq!(refs[0].cx, 5486400);
        assert_eq!(refs[0].cy, 3086100);
        assert_eq!(refs[0].data_rid, "rId5");
    }

    #[test]
    fn test_scan_smartart_refs_no_smartart() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
          <p:cSld><p:spTree>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>
              <p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody>
            </p:sp>
          </p:spTree></p:cSld>
        </p:sld>"#;

        let refs = scan_smartart_refs(slide_xml);
        assert!(refs.is_empty());
    }

    #[test]
    fn test_scan_smartart_refs_multiple() {
        let slide_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
               xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <p:cSld><p:spTree>
            <p:graphicFrame>
              <p:nvGraphicFramePr><p:cNvPr id="4" name="SmartArt1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
              <p:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></p:xfrm>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <dgm:relIds r:dm="rId10" r:lo="rId11" r:qs="rId12" r:cs="rId13"/>
              </a:graphicData></a:graphic>
            </p:graphicFrame>
            <p:graphicFrame>
              <p:nvGraphicFramePr><p:cNvPr id="5" name="SmartArt2"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
              <p:xfrm><a:off x="500" y="600"/><a:ext cx="700" cy="800"/></p:xfrm>
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                <dgm:relIds r:dm="rId20" r:lo="rId21" r:qs="rId22" r:cs="rId23"/>
              </a:graphicData></a:graphic>
            </p:graphicFrame>
          </p:spTree></p:cSld>
        </p:sld>"#;

        let refs = scan_smartart_refs(slide_xml);
        assert_eq!(refs.len(), 2);
        assert_eq!(refs[0].data_rid, "rId10");
        assert_eq!(refs[1].data_rid, "rId20");
        assert_eq!(refs[1].x, 500);
    }
}
