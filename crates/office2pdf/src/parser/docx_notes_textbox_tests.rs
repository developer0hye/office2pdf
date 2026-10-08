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

/// `document.xml` for one paragraph holding a `wp:inline` DrawingML text box
/// between the text `Before` and `After`.
///
/// Every dimension the tests assert on is a parameter — the `wp:extent`, the
/// `wps:spPr` children that carry `a:ln` and `a:solidFill`, and the
/// `wps:bodyPr` inset attributes — so each one can be varied and no measured
/// value can be satisfied by a hardcoded constant.
fn inline_text_box_document(
    width_emu: i64,
    height_emu: i64,
    shape_properties: &str,
    body_properties: &str,
) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
                        <wp:extent cx="{width_emu}" cy="{height_emu}"/>
                        <wp:docPr id="1" name="Text Box 1"/>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:cNvSpPr txBox="1"/>
                                    <wps:spPr>
                                        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                                        {shape_properties}
                                    </wps:spPr>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p>
                                                <w:r><w:t>Inside box</w:t></w:r>
                                            </w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr {body_properties}/>
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
</w:document>"#
    )
}

fn parse_flow_blocks(document_xml: &str) -> Vec<Block> {
    let data = build_docx_with_columns(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    match &doc.pages[0] {
        Page::Flow(flow) => flow.content.clone(),
        _ => panic!("Expected FlowPage"),
    }
}

/// Word draws a `wp:inline` text box as one item on the line of the paragraph
/// that anchors it, not as a paragraph of its own below it. Emitting the box's
/// paragraphs into the body gave the box's text its own line and moved every
/// line below the anchor: on the issue's package the anchor baseline read
/// 82.32pt against Word's 108.00pt (issue #1690).
#[test]
fn test_docx_inline_drawing_text_box_rides_its_anchor_paragraph() {
    let blocks = parse_flow_blocks(&inline_text_box_document(
        914_400,
        457_200,
        "",
        r#"lIns="91440" tIns="45720" rIns="91440" bIns="45720""#,
    ));

    let paragraphs: Vec<&Paragraph> = blocks
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect();
    assert_eq!(
        paragraphs.len(),
        1,
        "the box's text must not become a paragraph of its own, got: {:?}",
        paragraphs
            .iter()
            .map(|paragraph| paragraph
                .runs
                .iter()
                .map(|run| run.text.as_str())
                .collect::<String>())
            .collect::<Vec<String>>()
    );

    let runs = &paragraphs[0].runs;
    let texts: Vec<&str> = runs.iter().map(|run| run.text.as_str()).collect();
    assert_eq!(
        texts,
        vec!["Before", "", "After"],
        "the box's run sits between the runs it was written between"
    );
    assert!(
        runs[1].inline_box.is_some(),
        "the box rides the run at its own place in the paragraph"
    );

    let boxes: Vec<&InlineTextBox> = runs
        .iter()
        .filter_map(|run| run.inline_box.as_deref())
        .collect();
    assert_eq!(boxes.len(), 1, "exactly one run carries the box");
    let inline_box = boxes[0];
    assert_eq!(inline_box.width, 72.0, "914400 EMU is 72pt");
    assert_eq!(inline_box.height, 36.0, "457200 EMU is 36pt");
    let box_text: Vec<String> = inline_box
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => {
                Some(paragraph.runs.iter().map(|run| run.text.as_str()).collect())
            }
            _ => None,
        })
        .collect();
    assert_eq!(box_text, vec!["Inside box".to_string()]);
}

/// The extent, not a constant, gives the box its size.
#[test]
fn test_docx_inline_drawing_text_box_takes_its_own_extent() {
    for (width_emu, height_emu, width_pt, height_pt) in [
        (914_400_i64, 457_200_i64, 72.0, 36.0),
        (1_828_800, 274_320, 144.0, 21.6),
    ] {
        let blocks = parse_flow_blocks(&inline_text_box_document(
            width_emu,
            height_emu,
            "",
            r#"lIns="91440" tIns="45720" rIns="91440" bIns="45720""#,
        ));
        let inline_box = only_inline_text_box(&blocks);
        assert_eq!(inline_box.width, width_pt, "{width_emu} EMU");
        assert_eq!(inline_box.height, height_pt, "{height_emu} EMU");
    }
}

/// Word strokes the box's `a:ln` and fills its `a:solidFill`. Neither reached
/// the output at all before, so the box was invisible (issue #1690).
#[test]
fn test_docx_inline_drawing_text_box_paints_its_shape_frame() {
    for (line_width_emu, expected_pt) in [(9525_i64, 0.75), (6350, 0.5), (28_575, 2.25)] {
        let shape_properties = format!(
            r#"<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
               <a:ln w="{line_width_emu}"><a:solidFill><a:srgbClr val="203040"/></a:solidFill></a:ln>"#
        );
        let blocks = parse_flow_blocks(&inline_text_box_document(
            914_400,
            457_200,
            &shape_properties,
            r#"lIns="91440" tIns="45720" rIns="91440" bIns="45720""#,
        ));
        let inline_box = only_inline_text_box(&blocks);
        let stroke = inline_box
            .stroke
            .as_ref()
            .expect("an `a:ln` with a fill strokes the box");
        assert_eq!(stroke.width, expected_pt, "a:ln w={line_width_emu}");
        assert_eq!(
            stroke.color,
            Color {
                r: 0x20,
                g: 0x30,
                b: 0x40
            }
        );
        assert_eq!(
            inline_box.fill,
            Some(Color {
                r: 0xFF,
                g: 0xFF,
                b: 0xFF
            })
        );
    }
}

/// A box that declares no outline draws none.
#[test]
fn test_docx_inline_drawing_text_box_without_outline_draws_none() {
    let blocks = parse_flow_blocks(&inline_text_box_document(
        914_400,
        457_200,
        r#"<a:ln><a:noFill/></a:ln>"#,
        "",
    ));
    let inline_box = only_inline_text_box(&blocks);
    assert!(inline_box.stroke.is_none(), "`a:ln` with `a:noFill`");
    assert!(inline_box.fill.is_none(), "no `a:solidFill` on the shape");
}

/// `wps:bodyPr` states where the box's text starts. OOXML's defaults apply to
/// every side it leaves out, and to a box that declares no `wps:bodyPr`.
#[test]
fn test_docx_inline_drawing_text_box_insets() {
    let declared = only_inline_text_box(&parse_flow_blocks(&inline_text_box_document(
        914_400,
        457_200,
        "",
        r#"lIns="0" tIns="182880" rIns="45720" bIns="0""#,
    )));
    assert_eq!(declared.padding.left, 0.0);
    assert_eq!(declared.padding.top, 14.4, "182880 EMU is 14.4pt");
    assert_eq!(declared.padding.right, 3.6, "45720 EMU is 3.6pt");
    assert_eq!(declared.padding.bottom, 0.0);

    // No `wps:bodyPr` attributes at all: 91440 EMU sides, 45720 EMU top/bottom.
    let defaulted = only_inline_text_box(&parse_flow_blocks(&inline_text_box_document(
        914_400, 457_200, "", "",
    )));
    assert_eq!(defaulted.padding.left, 7.2);
    assert_eq!(defaulted.padding.right, 7.2);
    assert_eq!(defaulted.padding.top, 3.6);
    assert_eq!(defaulted.padding.bottom, 3.6);
}

/// A picture inside one of the box's paragraphs is its own `w:drawing`, with
/// its own `wp:extent` and `a:ln`. Only the outer `wps:wsp` describes the box,
/// so the nested drawing must not overwrite the box's size or frame — and must
/// not consume the box's slot in the scan, which would misalign every text box
/// after it.
#[test]
fn test_docx_inline_text_box_ignores_a_nested_drawing() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="1828800" cy="457200"/>
                        <wp:docPr id="1" name="Text Box 1"/>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:cNvSpPr txBox="1"/>
                                    <wps:spPr>
                                        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                                        <a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
                                    </wps:spPr>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p>
                                                <w:r><w:t>Caption</w:t></w:r>
                                                <w:r>
                                                    <w:drawing>
                                                        <wp:inline distT="0" distB="0" distL="0" distR="0">
                                                            <wp:extent cx="114300" cy="114300"/>
                                                            <wp:docPr id="2" name="Picture 2"/>
                                                            <a:graphic>
                                                                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                                                    <pic:pic>
                                                                        <pic:blipFill><a:blip r:embed="rIdMissing"/></pic:blipFill>
                                                                        <pic:spPr>
                                                                            <a:ln w="57150"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>
                                                                        </pic:spPr>
                                                                    </pic:pic>
                                                                </a:graphicData>
                                                            </a:graphic>
                                                        </wp:inline>
                                                    </w:drawing>
                                                </w:r>
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
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let inline_box = only_inline_text_box(&parse_flow_blocks(document_xml));
    assert_eq!(
        (inline_box.width, inline_box.height),
        (144.0, 36.0),
        "the box keeps its own 1828800 x 457200 EMU extent, not the picture's 114300"
    );
    let stroke = inline_box.stroke.as_ref().expect("the box's own `a:ln`");
    assert_eq!(
        stroke.width, 0.75,
        "the box keeps its own 9525 EMU outline, not the picture's 57150"
    );
    assert_eq!(stroke.color, Color { r: 0, g: 0, b: 0 });
}

/// An anchored box is positioned by its own offsets and keeps that path. Only
/// `wp:inline` joins the anchor paragraph's line.
#[test]
fn test_docx_anchored_drawing_text_box_stays_floating() {
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
                    <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
                               relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1"
                               allowOverlap="1">
                        <wp:simplePos x="0" y="0"/>
                        <wp:positionH relativeFrom="column"><wp:posOffset>228600</wp:posOffset></wp:positionH>
                        <wp:positionV relativeFrom="paragraph"><wp:posOffset>114300</wp:posOffset></wp:positionV>
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:wrapNone/>
                        <wp:docPr id="1" name="Text Box 1"/>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:cNvSpPr txBox="1"/>
                                    <wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p><w:r><w:t>Anchored box</w:t></w:r></w:p>
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

    let blocks = parse_flow_blocks(document_xml);
    let floating: Vec<&FloatingTextBox> = blocks
        .iter()
        .filter_map(|block| match block {
            Block::FloatingTextBox(text_box) => Some(text_box),
            _ => None,
        })
        .collect();
    assert_eq!(floating.len(), 1, "got blocks: {blocks:?}");
    assert_eq!(floating[0].offset_x, 18.0, "228600 EMU is 18pt");
    assert_eq!(floating[0].offset_y, 9.0, "114300 EMU is 9pt");
    assert!(
        blocks
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => Some(paragraph),
                _ => None,
            })
            .all(|paragraph| paragraph.runs.iter().all(|run| run.inline_box.is_none())),
        "an anchored box never rides a run"
    );
}

/// The single inline text box the blocks carry, cloned so a caller can pass a
/// temporary `parse_flow_blocks` result straight in.
fn only_inline_text_box(blocks: &[Block]) -> InlineTextBox {
    let mut found: Vec<&InlineTextBox> = Vec::new();
    for block in blocks {
        if let Block::Paragraph(paragraph) = block {
            found.extend(
                paragraph
                    .runs
                    .iter()
                    .filter_map(|run| run.inline_box.as_deref()),
            );
        }
    }
    assert_eq!(found.len(), 1, "expected exactly one inline text box");
    found[0].clone()
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

    let blocks = parse_flow_blocks(document_xml);

    // The body keeps only its own three paragraphs: the lead-in, the one that
    // anchors the box, and the tail. The box's two paragraphs stay inside the
    // box (issue #1690).
    let body_texts: Vec<String> = blocks
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => {
                Some(paragraph.runs.iter().map(|run| run.text.as_str()).collect())
            }
            _ => None,
        })
        .collect();
    assert_eq!(
        body_texts,
        vec!["Lead-in".to_string(), String::new(), "Tail".to_string()]
    );

    let inline_box = only_inline_text_box(&blocks);
    let box_texts: Vec<String> = inline_box
        .content
        .iter()
        .map(|paragraph| paragraph.runs.iter().map(|run| run.text.as_str()).collect())
        .collect();
    assert_eq!(
        box_texts,
        vec!["First line".to_string(), "Second line".to_string()],
        "both of the box's paragraphs stay in the box, in order"
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

    let inline_text_box = flow.content.iter().find_map(|block| match block {
        Block::Paragraph(paragraph) => paragraph
            .runs
            .iter()
            .find_map(|run| run.inline_box.as_deref()),
        _ => None,
    });
    assert!(
        inline_text_box.is_some(),
        "an inline text box holding a table stays attached to its anchor run"
    );
    assert!(
        !flow
            .content
            .iter()
            .any(|block| matches!(block, Block::Table(_))),
        "a table inside an inline text box must not flatten into body flow"
    );

    let table = inline_text_box
        .expect("the inline box must remain attached to its anchor run")
        .content
        .iter()
        .find_map(|block| match block {
            Block::Table(table) => Some(table),
            _ => None,
        })
        .expect("the box keeps its table in its own flow");
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
fn test_docx_inline_drawing_text_box_keeps_picture_in_its_flow() {
    let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p><w:r><w:drawing>
            <wp:inline distT="0" distB="0" distL="0" distR="0">
                <wp:extent cx="914400" cy="457200"/>
                <wp:docPr id="1" name="Text Box 1"/>
                <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                    <wps:wsp>
                        <wps:cNvSpPr txBox="1"/>
                        <wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>
                        <wps:txbx><w:txbxContent>
                            <w:p><w:r><w:drawing>
                                <wp:inline distT="0" distB="0" distL="0" distR="0">
                                    <wp:extent cx="457200" cy="457200"/>
                                    <wp:docPr id="2" name="Picture 1"/>
                                    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                        <pic:pic>
                                            <pic:nvPicPr><pic:cNvPr id="2" name="image1.bmp"/><pic:cNvPicPr/></pic:nvPicPr>
                                            <pic:blipFill><a:blip r:embed="rIdImage1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                                            <pic:spPr>
                                                <a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>
                                                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                                            </pic:spPr>
                                        </pic:pic>
                                    </a:graphicData></a:graphic>
                                </wp:inline>
                            </w:drawing></w:r></w:p>
                        </w:txbxContent></wps:txbx>
                        <wps:bodyPr/>
                    </wps:wsp>
                </a:graphicData></a:graphic>
            </wp:inline>
        </w:drawing></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

    let data = super::image_tests::build_docx_with_custom_image_document(document_xml);
    let parser = DocxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
    let flow = match &doc.pages[0] {
        Page::Flow(flow) => flow,
        _ => panic!("Expected FlowPage"),
    };
    let inline_box = flow.content.iter().find_map(|block| match block {
        Block::Paragraph(paragraph) => paragraph
            .runs
            .iter()
            .find_map(|run| run.inline_box.as_deref()),
        _ => None,
    });
    let image = inline_box
        .expect("the picture's text box stays attached to its anchor run")
        .content
        .iter()
        .find_map(|block| match block {
            Block::Image(image) => Some(image),
            _ => None,
        })
        .expect("the text box keeps the picture in its own flow");
    assert!(!image.data.is_empty());
    assert!(
        !flow
            .content
            .iter()
            .any(|block| matches!(block, Block::Image(_))),
        "a picture inside an inline box must not flatten into body flow"
    );
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
fn test_docx_anchored_text_box_cell_nil_borders_hide_table_grid() {
    let data: &[u8] = include_bytes!("../../../../tests/fixtures/docx/libreoffice/tdf105688.docx");
    let (document, _warnings) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();

    let text_box: &FloatingTextBox = document
        .pages
        .iter()
        .find_map(|page| {
            let content = match page {
                Page::Flow(page) | Page::FlowContinuous(page) => &page.content,
                _ => return None,
            };
            content.iter().find_map(|block| match block {
                Block::FloatingTextBox(text_box) => Some(text_box),
                _ => None,
            })
        })
        .expect("the fixture contains a page-anchored text box");
    let table = text_box
        .content
        .iter()
        .find_map(|block| match block {
            Block::Table(table) => Some(table),
            _ => None,
        })
        .expect("the floating text box contains its pull-quote table");

    assert_eq!(table.rows.len(), 1);
    assert_eq!(table.rows[0].cells.len(), 1);
    assert!(
        table.rows[0].cells[0].border.is_none(),
        "the cell's four explicit tcBorders nil values suppress the table grid"
    );
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
