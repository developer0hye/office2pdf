use super::*;

// ── Helpers ──────────────────────────────────────────────────────────

/// A placeholder `<p:sp>` with no explicit geometry and no local styling:
/// everything must come from the inheritance chain.
fn make_plain_placeholder_sp(ph_attrs: &str, paragraphs_xml: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph {ph_attrs}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{paragraphs_xml}</p:txBody></p:sp>"#
    )
}

fn make_simple_paragraph(text: &str) -> String {
    format!(r#"<a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p>"#)
}

fn make_leveled_paragraph(level: u32, text: &str) -> String {
    format!(r#"<a:p><a:pPr lvl="{level}"/><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p>"#)
}

fn make_slide(shapes: &[String]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
    );
    for shape in shapes {
        xml.push_str(shape);
    }
    xml.push_str("</p:spTree></p:cSld></p:sld>");
    xml
}

fn make_layout(shapes: &[String]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
    );
    for shape in shapes {
        xml.push_str(shape);
    }
    xml.push_str("</p:spTree></p:cSld></p:sldLayout>");
    xml
}

/// Master XML with a `<p:txStyles>` block after `cSld`.
fn make_master_with_tx_styles(tx_styles_xml: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:txStyles>{tx_styles_xml}</p:txStyles></p:sldMaster>"#
    )
}

fn parse_document(data: &[u8]) -> Document {
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(data, &ConvertOptions::default()).unwrap();
    doc
}

/// Collect every (text, style, paragraph_style) triple on the first page,
/// from both Paragraph and List blocks.
fn collect_runs(doc: &Document) -> Vec<(String, TextStyle, ParagraphStyle)> {
    let page = first_fixed_page(doc);
    let mut out: Vec<(String, TextStyle, ParagraphStyle)> = Vec::new();
    for element in &page.elements {
        let FixedElementKind::TextBox(text_box) = &element.kind else {
            continue;
        };
        let mut visit_paragraph = |paragraph: &Paragraph| {
            for run in &paragraph.runs {
                out.push((run.text.clone(), run.style.clone(), paragraph.style.clone()));
            }
        };
        for block in &text_box.content {
            match block {
                Block::Paragraph(paragraph) => visit_paragraph(paragraph),
                Block::List(list) => {
                    for item in &list.items {
                        for paragraph in &item.content {
                            visit_paragraph(paragraph);
                        }
                    }
                }
                _ => {}
            }
        }
    }
    out
}

fn run_for<'a>(
    runs: &'a [(String, TextStyle, ParagraphStyle)],
    needle: &str,
) -> &'a (String, TextStyle, ParagraphStyle) {
    runs.iter()
        .find(|(text, _, _)| text.contains(needle))
        .unwrap_or_else(|| panic!("no run containing {needle:?}"))
}

const TITLE_STYLE_44_CTR: &str =
    r#"<p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>"#;

// ── Master txStyles inheritance ──────────────────────────────────────

#[test]
fn test_title_placeholder_inherits_master_title_style() {
    let slide = make_slide(&[make_plain_placeholder_sp(
        r#"type="title""#,
        &make_simple_paragraph("Big Title"),
    )]);
    let layout = make_layout(&[]);
    let master = make_master_with_tx_styles(TITLE_STYLE_44_CTR);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let runs = collect_runs(&doc);
    let (_, style, paragraph_style) = run_for(&runs, "Big Title");
    assert_eq!(style.font_size, Some(44.0));
    assert_eq!(paragraph_style.alignment, Some(Alignment::Center));
}

#[test]
fn test_ctr_title_placeholder_inherits_master_title_style() {
    let slide = make_slide(&[make_plain_placeholder_sp(
        r#"type="ctrTitle""#,
        &make_simple_paragraph("Cover Title"),
    )]);
    let layout = make_layout(&[]);
    let master = make_master_with_tx_styles(TITLE_STYLE_44_CTR);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let runs = collect_runs(&doc);
    let (_, style, _) = run_for(&runs, "Cover Title");
    assert_eq!(style.font_size, Some(44.0));
}

#[test]
fn test_body_placeholder_inherits_master_body_style_levels() {
    let body_style = r#"<p:bodyStyle><a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr><a:lvl2pPr><a:defRPr sz="2400"/></a:lvl2pPr></p:bodyStyle>"#;
    let paragraphs = format!(
        "{}{}",
        make_simple_paragraph("Level one"),
        make_leveled_paragraph(1, "Level two")
    );
    let slide = make_slide(&[make_plain_placeholder_sp(
        r#"type="body" idx="1""#,
        &paragraphs,
    )]);
    let layout = make_layout(&[]);
    let master = make_master_with_tx_styles(body_style);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let runs = collect_runs(&doc);
    assert_eq!(run_for(&runs, "Level one").1.font_size, Some(28.0));
    assert_eq!(run_for(&runs, "Level two").1.font_size, Some(24.0));
}

#[test]
fn test_footer_placeholder_inherits_master_other_style() {
    let other_style =
        r#"<p:otherStyle><a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr></p:otherStyle>"#;
    let slide = make_slide(&[make_plain_placeholder_sp(
        r#"type="ftr" sz="quarter" idx="11""#,
        &make_simple_paragraph("Footer text"),
    )]);
    let layout = make_layout(&[]);
    let master = make_master_with_tx_styles(other_style);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let runs = collect_runs(&doc);
    assert_eq!(run_for(&runs, "Footer text").1.font_size, Some(12.0));
}

// ── Layout placeholder lstStyle overrides ────────────────────────────

#[test]
fn test_layout_placeholder_list_style_overrides_master() {
    let layout_title = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="3000"/></a:lvl1pPr></a:lstStyle><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>"#;
    let slide = make_slide(&[make_plain_placeholder_sp(
        r#"type="title""#,
        &make_simple_paragraph("Layout styled"),
    )]);
    let layout = make_layout(&[layout_title.to_string()]);
    let master = make_master_with_tx_styles(TITLE_STYLE_44_CTR);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let runs = collect_runs(&doc);
    let (_, style, paragraph_style) = run_for(&runs, "Layout styled");
    // Layout lstStyle wins over master titleStyle...
    assert_eq!(style.font_size, Some(30.0));
    assert_eq!(paragraph_style.alignment, Some(Alignment::Left));
}

// ── Theme font resolution ────────────────────────────────────────────

#[test]
fn test_title_style_resolves_theme_major_font() {
    let title_style = r#"<p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle>"#;
    let slide = make_slide(&[make_plain_placeholder_sp(
        r#"type="title""#,
        &make_simple_paragraph("Themed Title"),
    )]);
    let layout = make_layout(&[]);
    let master = make_master_with_tx_styles(title_style);
    let theme_xml = make_theme_xml(
        &standard_theme_colors(),
        "Test Major Font",
        "Test Minor Font",
    );
    let data = build_test_pptx_with_theme_layout_master(
        SLIDE_CX, SLIDE_CY, &slide, &layout, &master, &theme_xml,
    );

    let doc = parse_document(&data);
    let runs = collect_runs(&doc);
    let (_, style, _) = run_for(&runs, "Themed Title");
    assert_eq!(style.font_family.as_deref(), Some("Test Major Font"));
}

// ── Precedence guards ────────────────────────────────────────────────

#[test]
fn test_explicit_run_properties_override_inherited_style() {
    let paragraph =
        r#"<a:p><a:r><a:rPr lang="en-US" sz="2000"/><a:t>Explicit size</a:t></a:r></a:p>"#;
    let slide = make_slide(&[make_plain_placeholder_sp(r#"type="title""#, paragraph)]);
    let layout = make_layout(&[]);
    let master = make_master_with_tx_styles(TITLE_STYLE_44_CTR);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let runs = collect_runs(&doc);
    assert_eq!(run_for(&runs, "Explicit size").1.font_size, Some(20.0));
}

#[test]
fn test_non_placeholder_text_box_ignores_title_style() {
    let text_box = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="5000000" cy="500000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Plain box</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let slide = make_slide(&[text_box.to_string()]);
    let layout = make_layout(&[]);
    let master = make_master_with_tx_styles(TITLE_STYLE_44_CTR);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let runs = collect_runs(&doc);
    assert_ne!(run_for(&runs, "Plain box").1.font_size, Some(44.0));
}

#[test]
fn test_run_typeface_overrides_inherited_default_font() {
    // A run's own <a:latin> must beat the family inherited from the master
    // otherStyle; the old first-wins guard kept the inherited font and
    // dropped explicit run fonts (CJK runs then lost their bold variants).
    let other_style = r#"<p:otherStyle><a:defPPr/><a:lvl1pPr><a:defRPr><a:latin typeface="Calibri"/></a:defRPr></a:lvl1pPr></p:otherStyle>"#;
    let text_box = r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="5000000" cy="500000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" b="1"><a:latin typeface="Malgun Gothic"/><a:ea typeface="Malgun Gothic"/></a:rPr><a:t>Own font</a:t></a:r></a:p></p:txBody></p:sp>"#;
    let slide = make_slide(&[text_box.to_string()]);
    let layout = make_layout(&[]);
    let master = make_master_with_tx_styles(other_style);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let runs = collect_runs(&doc);
    let (_, style, _) = run_for(&runs, "Own font");
    assert_eq!(style.font_family.as_deref(), Some("Malgun Gothic"));
    assert_eq!(style.bold, Some(true));
}
