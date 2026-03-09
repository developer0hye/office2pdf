use super::*;
#[path = "docx_notes_toc_tests.rs"]
mod notes_toc_tests;

#[test]
fn test_docx_sdt_with_paragraphs() {
    let sdt = docx_rs::StructuredDataTag::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("SDT Content")),
    );

    let docx = docx_rs::Docx::new().add_structured_data_tag(sdt);

    let buf = Vec::new();
    let mut cursor = Cursor::new(buf);
    docx.build().pack(&mut cursor).unwrap();
    let data = cursor.into_inner();

    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let page = &doc.pages[0];
    let content = match page {
        Page::Flow(fp) => &fp.content,
        _ => panic!("Expected FlowPage"),
    };

    let all_text: Vec<String> = content
        .iter()
        .filter_map(|b| match b {
            Block::Paragraph(p) => {
                let t: String = p.runs.iter().map(|r| r.text.clone()).collect();
                if t.is_empty() { None } else { Some(t) }
            }
            _ => None,
        })
        .collect();

    assert!(
        all_text.iter().any(|t| t.contains("SDT Content")),
        "Expected 'SDT Content' in output, got: {all_text:?}"
    );
}

#[test]
fn test_docx_drawing_text_box_paragraph_is_emitted() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p>
            <w:r><w:t>Before</w:t></w:r>
            <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:docPr id="1" name="Text Box 1"/>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p>
                                                <w:r><w:t>Inside box</w:t></w:r>
                                            </w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
            <w:r><w:t>After</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let texts: Vec<String> = match &doc.pages[0] {
        Page::Flow(flow) => flow
            .content
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                _ => None,
            })
            .collect(),
        _ => panic!("Expected FlowPage"),
    };

    assert_eq!(
        texts,
        vec![
            "Before".to_string(),
            "Inside box".to_string(),
            "After".to_string(),
        ]
    );
}

#[test]
fn test_docx_drawing_text_box_multiple_paragraphs_are_emitted_in_order() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p><w:r><w:t>Lead-in</w:t></w:r></w:p>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:docPr id="1" name="Text Box 2"/>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p><w:r><w:t>First line</w:t></w:r></w:p>
                                            <w:p><w:r><w:t>Second line</w:t></w:r></w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
        <w:p><w:r><w:t>Tail</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let texts: Vec<String> = match &doc.pages[0] {
        Page::Flow(flow) => flow
            .content
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                _ => None,
            })
            .collect(),
        _ => panic!("Expected FlowPage"),
    };

    assert_eq!(
        texts,
        vec![
            "Lead-in".to_string(),
            "First line".to_string(),
            "Second line".to_string(),
            "Tail".to_string(),
        ]
    );
}

#[test]
fn test_docx_drawing_text_box_table_is_emitted() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p><w:r><w:t>Before table box</w:t></w:r></w:p>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:docPr id="1" name="Text Box Table"/>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:tbl>
                                                <w:tblPr/>
                                                <w:tblGrid>
                                                    <w:gridCol w:w="2000"/>
                                                    <w:gridCol w:w="2000"/>
                                                </w:tblGrid>
                                                <w:tr>
                                                    <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
                                                    <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
                                                </w:tr>
                                            </w:tbl>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
        <w:p><w:r><w:t>After table box</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };

    let has_table = flow
        .content
        .iter()
        .any(|block| matches!(block, Block::Table(_)));
    assert!(has_table, "Expected a table extracted from text box");

    let table = first_table(&doc);
    assert_eq!(table.rows.len(), 1);
    assert_eq!(table.rows[0].cells.len(), 2);

    let cell_text: Vec<String> = table.rows[0]
        .cells
        .iter()
        .map(|cell| {
            cell.content
                .iter()
                .filter_map(|block| match block {
                    Block::Paragraph(p) => Some(
                        p.runs
                            .iter()
                            .map(|run| run.text.as_str())
                            .collect::<String>(),
                    ),
                    _ => None,
                })
                .collect::<String>()
        })
        .collect();
    assert_eq!(cell_text, vec!["A".to_string(), "B".to_string()]);
}

#[test]
fn test_docx_vml_text_box_paragraph_is_emitted() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml">
    <w:body>
        <w:p>
            <w:r><w:t>Before</w:t></w:r>
            <w:r>
                <w:pict>
                    <v:shape id="TextBox1" style="width:100pt;height:40pt">
                        <v:textbox>
                            <w:txbxContent>
                                <w:p><w:r><w:t>VML box</w:t></w:r></w:p>
                            </w:txbxContent>
                        </v:textbox>
                    </v:shape>
                </w:pict>
            </w:r>
            <w:r><w:t>After</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let texts: Vec<String> = match &doc.pages[0] {
        Page::Flow(flow) => flow
            .content
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                _ => None,
            })
            .collect(),
        _ => panic!("Expected FlowPage"),
    };

    assert_eq!(
        texts,
        vec![
            "Before".to_string(),
            "VML box".to_string(),
            "After".to_string(),
        ]
    );
}

#[test]
fn test_docx_vml_text_box_multiple_paragraphs_are_emitted_in_order() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml">
    <w:body>
        <w:p><w:r><w:t>Lead-in</w:t></w:r></w:p>
        <w:p>
            <w:r>
                <w:pict>
                    <v:shape id="TextBox2" style="width:120pt;height:60pt">
                        <v:textbox>
                            <w:txbxContent>
                                <w:p><w:r><w:t>First VML line</w:t></w:r></w:p>
                                <w:p><w:r><w:t>Second VML line</w:t></w:r></w:p>
                            </w:txbxContent>
                        </v:textbox>
                    </v:shape>
                </w:pict>
            </w:r>
        </w:p>
        <w:p><w:r><w:t>Tail</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let texts: Vec<String> = match &doc.pages[0] {
        Page::Flow(flow) => flow
            .content
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                _ => None,
            })
            .collect(),
        _ => panic!("Expected FlowPage"),
    };

    assert_eq!(
        texts,
        vec![
            "Lead-in".to_string(),
            "First VML line".to_string(),
            "Second VML line".to_string(),
            "Tail".to_string(),
        ]
    );
}

#[test]
fn test_docx_vml_floating_text_box_square_wrap() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:w10="urn:schemas-microsoft-com:office:word">
    <w:body>
        <w:p>
            <w:r><w:t>Before</w:t></w:r>
            <w:r>
                <w:pict>
                    <v:shape id="TextBox3"
                             style="position:absolute;margin-left:72pt;margin-top:36pt;width:144pt;height:72pt;z-index:1;visibility:visible;mso-wrap-style:square">
                        <v:textbox>
                            <w:txbxContent>
                                <w:p><w:r><w:t>VML floating box</w:t></w:r></w:p>
                            </w:txbxContent>
                        </v:textbox>
                    </v:shape>
                    <w10:wrap type="square"/>
                </w:pict>
            </w:r>
            <w:r><w:t>After</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_text_boxes(&doc);
    assert_eq!(floating.len(), 1, "Expected one floating VML text box");

    let ftb = floating[0];
    assert_eq!(ftb.wrap_mode, WrapMode::Square);
    assert!((ftb.offset_x - 72.0).abs() < 0.5);
    assert!((ftb.offset_y - 36.0).abs() < 0.5);
    assert!((ftb.width - 144.0).abs() < 0.5);
    assert!((ftb.height - 72.0).abs() < 0.5);

    let texts: Vec<String> = ftb
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
            _ => None,
        })
        .collect();
    assert_eq!(texts, vec!["VML floating box".to_string()]);
}

#[test]
fn test_docx_vml_floating_text_box_none_wrap() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:w10="urn:schemas-microsoft-com:office:word">
    <w:body>
        <w:p>
            <w:r>
                <w:pict>
                    <v:shape id="TextBox4"
                             style="position:absolute;margin-left:12pt;margin-top:18pt;width:90pt;height:40pt;z-index:1;visibility:visible;mso-wrap-style:square">
                        <v:textbox>
                            <w:txbxContent>
                                <w:p><w:r><w:t>No wrap box</w:t></w:r></w:p>
                            </w:txbxContent>
                        </v:textbox>
                    </v:shape>
                    <w10:wrap type="none"/>
                </w:pict>
            </w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_text_boxes(&doc);
    assert_eq!(floating.len(), 1, "Expected one floating VML text box");
    assert_eq!(floating[0].wrap_mode, WrapMode::None);
}

fn find_floating_text_boxes(doc: &Document) -> Vec<&FloatingTextBox> {
    let page = match &doc.pages[0] {
        Page::Flow(f) => f,
        _ => panic!("Expected FlowPage"),
    };
    page.content
        .iter()
        .filter_map(|b| match b {
            Block::FloatingTextBox(ftb) => Some(ftb),
            _ => None,
        })
        .collect()
}

#[test]
fn test_docx_floating_text_box_square_wrap() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p>
            <w:r><w:t>Before</w:t></w:r>
            <w:r>
                <w:drawing>
                    <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" allowOverlap="0" behindDoc="0" locked="0" layoutInCell="1" relativeHeight="251659264">
                        <wp:simplePos x="0" y="0"/>
                        <wp:positionH relativeFrom="margin"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                        <wp:positionV relativeFrom="margin"><wp:posOffset>457200</wp:posOffset></wp:positionV>
                        <wp:extent cx="1828800" cy="914400"/>
                        <wp:effectExtent l="0" t="0" r="0" b="0"/>
                        <wp:wrapSquare wrapText="bothSides"/>
                        <wp:docPr id="1" name="Anchored Text Box"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p><w:r><w:t>Inside anchored box</w:t></w:r></w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:anchor>
                </w:drawing>
            </w:r>
            <w:r><w:t>After</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_text_boxes(&doc);
    assert_eq!(floating.len(), 1, "Expected one floating text box");

    let ftb = floating[0];
    assert_eq!(ftb.wrap_mode, WrapMode::Square);
    assert!(
        (ftb.offset_x - 72.0).abs() < 0.5,
        "Expected offset_x ~72pt, got {}",
        ftb.offset_x
    );
    assert!(
        (ftb.offset_y - 36.0).abs() < 0.5,
        "Expected offset_y ~36pt, got {}",
        ftb.offset_y
    );
    assert!(
        (ftb.width - 144.0).abs() < 0.5,
        "Expected width ~144pt, got {}",
        ftb.width
    );
    assert!(
        (ftb.height - 72.0).abs() < 0.5,
        "Expected height ~72pt, got {}",
        ftb.height
    );

    let texts: Vec<String> = ftb
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
            _ => None,
        })
        .collect();
    assert_eq!(texts, vec!["Inside anchored box".to_string()]);
}

#[test]
fn test_docx_floating_text_box_top_and_bottom_wrap() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" allowOverlap="1" behindDoc="0" locked="0" layoutInCell="1" relativeHeight="251659264">
                        <wp:simplePos x="0" y="0"/>
                        <wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>
                        <wp:positionV relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionV>
                        <wp:extent cx="1270000" cy="635000"/>
                        <wp:effectExtent l="0" t="0" r="0" b="0"/>
                        <wp:wrapTopAndBottom/>
                        <wp:docPr id="2" name="Top Bottom Text Box"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p><w:r><w:t>Top and bottom box</w:t></w:r></w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:anchor>
                </w:drawing>
            </w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let floating = find_floating_text_boxes(&doc);
    assert_eq!(floating.len(), 1, "Expected one floating text box");
    assert_eq!(floating[0].wrap_mode, WrapMode::TopAndBottom);
}
