use super::*;

/// Helper to extract texts from SmartArtNode list.
fn texts(nodes: &[SmartArtNode]) -> Vec<&str> {
    nodes.iter().map(|n| n.text.as_str()).collect()
}

/// Helper to extract depths from SmartArtNode list.
fn depths(nodes: &[SmartArtNode]) -> Vec<usize> {
    nodes.iter().map(|n| n.depth).collect()
}

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
          <dgm:cxnLst>
            <dgm:cxn modelId="10" type="parOf" srcId="0" destId="1"/>
            <dgm:cxn modelId="11" type="parOf" srcId="0" destId="2"/>
            <dgm:cxn modelId="12" type="parOf" srcId="0" destId="3"/>
          </dgm:cxnLst>
        </dgm:dataModel>"#;

    let items = parse_smartart_data_xml(xml);
    assert_eq!(texts(&items), vec!["Step 1", "Step 2", "Step 3"]);
    // All direct children of doc -> depth 0
    assert_eq!(depths(&items), vec![0, 0, 0]);
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
          <dgm:cxnLst>
            <dgm:cxn modelId="20" type="parOf" srcId="0" destId="1"/>
            <dgm:cxn modelId="21" type="parOf" srcId="0" destId="2"/>
          </dgm:cxnLst>
        </dgm:dataModel>"#;

    let items = parse_smartart_data_xml(xml);
    assert_eq!(texts(&items), vec!["Item A", "Item B"]);
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
    assert_eq!(texts(&items), vec!["Valid"]);
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
    assert_eq!(texts(&items), vec!["Hello World"]);
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
    assert_eq!(texts(&items), vec!["Implicit Node"]);
}

#[test]
fn test_parse_smartart_data_with_hierarchy() {
    // Hierarchy: doc -> A, B; A -> C (so C is depth 1)
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dgm:ptLst>
            <dgm:pt modelId="0" type="doc">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Root</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="1" type="node">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Manager A</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="2" type="node">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Manager B</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="3" type="node">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Employee C</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
          </dgm:ptLst>
          <dgm:cxnLst>
            <dgm:cxn modelId="10" type="parOf" srcId="0" destId="1"/>
            <dgm:cxn modelId="11" type="parOf" srcId="0" destId="2"/>
            <dgm:cxn modelId="12" type="parOf" srcId="1" destId="3"/>
          </dgm:cxnLst>
        </dgm:dataModel>"#;

    let items = parse_smartart_data_xml(xml);
    assert_eq!(texts(&items), vec!["Manager A", "Manager B", "Employee C"]);
    assert_eq!(depths(&items), vec![0, 0, 1]);
}

#[test]
fn test_parse_smartart_data_deep_hierarchy() {
    // doc -> A -> B -> C (depths 0, 1, 2)
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dgm:ptLst>
            <dgm:pt modelId="0" type="doc">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Root</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="1" type="node">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Level 0</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="2" type="node">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Level 1</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
            <dgm:pt modelId="3" type="node">
              <dgm:t><a:bodyPr/><a:p><a:r><a:t>Level 2</a:t></a:r></a:p></dgm:t>
            </dgm:pt>
          </dgm:ptLst>
          <dgm:cxnLst>
            <dgm:cxn modelId="10" type="parOf" srcId="0" destId="1"/>
            <dgm:cxn modelId="11" type="parOf" srcId="1" destId="2"/>
            <dgm:cxn modelId="12" type="parOf" srcId="2" destId="3"/>
          </dgm:cxnLst>
        </dgm:dataModel>"#;

    let items = parse_smartart_data_xml(xml);
    assert_eq!(texts(&items), vec!["Level 0", "Level 1", "Level 2"]);
    assert_eq!(depths(&items), vec![0, 1, 2]);
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
