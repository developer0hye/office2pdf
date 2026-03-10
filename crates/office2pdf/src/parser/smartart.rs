use quick_xml::Reader;
/// SmartArt diagram parser for PPTX files.
///
/// Parses SmartArt data model XML to extract text content from diagram nodes.
/// SmartArt diagrams use the DrawingML Diagram namespace
/// (`http://schemas.openxmlformats.org/drawingml/2006/diagram`).
use quick_xml::events::Event;

use crate::ir::SmartArtNode;
use std::collections::HashMap;

/// Internal representation of a parsed SmartArt point before depth resolution.
struct RawPoint {
    model_id: String,
    pt_type: String,
    text: String,
}

/// Parse SmartArt data model XML and extract nodes with hierarchy depth.
///
/// The data model XML contains `<dgm:ptLst>` with `<dgm:pt>` elements and
/// `<dgm:cxnLst>` with `<dgm:cxn>` connections (type="parOf" links parent→child).
/// We extract text from `type="node"` points, then compute depth from the
/// connection graph rooted at the `type="doc"` node.
pub(crate) fn parse_smartart_data_xml(xml: &str) -> Vec<SmartArtNode> {
    let (points, connections) = parse_points_and_connections(xml);

    // Separate doc root and node points
    let mut doc_id: Option<String> = None;
    let mut node_texts: HashMap<String, String> = HashMap::new();
    let mut node_order: Vec<String> = Vec::new();

    for pt in &points {
        match pt.pt_type.as_str() {
            "doc" => {
                doc_id = Some(pt.model_id.clone());
            }
            "node" => {
                if !pt.text.is_empty() {
                    node_texts.insert(pt.model_id.clone(), pt.text.clone());
                    node_order.push(pt.model_id.clone());
                }
            }
            _ => {}
        }
    }

    // Build parent→children map from "parOf" connections
    let mut children_map: HashMap<String, Vec<String>> = HashMap::new();
    for (src, dest) in &connections {
        children_map
            .entry(src.clone())
            .or_default()
            .push(dest.clone());
    }

    // BFS from doc root to assign depth to each node
    let mut depth_map: HashMap<String, usize> = HashMap::new();
    if let Some(root_id) = doc_id {
        let mut queue = std::collections::VecDeque::new();
        // Children of doc root are depth 0
        if let Some(children) = children_map.get(&root_id) {
            for child in children {
                queue.push_back((child.clone(), 0usize));
            }
        }
        while let Some((id, depth)) = queue.pop_front() {
            if depth_map.contains_key(&id) {
                continue;
            }
            depth_map.insert(id.clone(), depth);
            if let Some(children) = children_map.get(&id) {
                for child in children {
                    if !depth_map.contains_key(child) {
                        queue.push_back((child.clone(), depth + 1));
                    }
                }
            }
        }
    }

    // Build result in original document order, using depth from BFS
    node_order
        .iter()
        .filter_map(|id| {
            node_texts.get(id).map(|text| SmartArtNode {
                text: text.clone(),
                depth: depth_map.get(id).copied().unwrap_or(0),
            })
        })
        .collect()
}

/// Parse both `<dgm:ptLst>` points and `<dgm:cxnLst>` connections from SmartArt XML.
///
/// Returns (points, connections) where connections are (srcId, destId) pairs
/// for "parOf" type connections.
fn parse_points_and_connections(xml: &str) -> (Vec<RawPoint>, Vec<(String, String)>) {
    let mut points = Vec::new();
    let mut connections = Vec::new();
    let mut reader = Reader::from_str(xml);

    // Point parsing state
    let mut in_pt = false;
    let mut pt_model_id = String::new();
    let mut pt_type = String::new();
    let mut in_t_block = false;
    let mut in_a_r = false;
    let mut in_a_t = false;
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
                        pt_model_id.clear();
                        pt_type = String::from("node");

                        for attr in e.attributes().flatten() {
                            if let Ok(v) = attr.unescape_value() {
                                match attr.key.local_name().as_ref() {
                                    b"type" => pt_type = v.to_string(),
                                    b"modelId" => pt_model_id = v.to_string(),
                                    _ => {}
                                }
                            }
                        }
                    }
                    b"pt" if in_pt => {
                        pt_depth += 1;
                    }
                    b"t" if in_a_r => {
                        in_a_t = true;
                    }
                    b"r" if in_t_block => {
                        in_a_r = true;
                    }
                    b"t" if in_pt && !in_t_block => {
                        in_t_block = true;
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"cxn" {
                    // <dgm:cxn modelId="N" type="parOf" srcId="S" destId="D"/>
                    let mut cxn_type = String::new();
                    let mut src_id = String::new();
                    let mut dest_id = String::new();
                    for attr in e.attributes().flatten() {
                        if let Ok(v) = attr.unescape_value() {
                            match attr.key.local_name().as_ref() {
                                b"type" => cxn_type = v.to_string(),
                                b"srcId" => src_id = v.to_string(),
                                b"destId" => dest_id = v.to_string(),
                                _ => {}
                            }
                        }
                    }
                    if cxn_type == "parOf" && !src_id.is_empty() && !dest_id.is_empty() {
                        connections.push((src_id, dest_id));
                    }
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
                            let trimmed = current_text.trim().to_string();
                            points.push(RawPoint {
                                model_id: pt_model_id.clone(),
                                pt_type: pt_type.clone(),
                                text: trimmed,
                            });
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

    (points, connections)
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
#[path = "smartart_tests.rs"]
mod tests;
