use std::collections::HashMap;

use super::*;
use crate::ir::{Color, LineCap, LineJoin};
use crate::parser::drawingml::ThemeFontScheme;

fn accent_theme() -> HashMap<String, Color> {
    // Office default theme slots a workbook actually ships in theme1.xml.
    HashMap::from([
        ("dk1".to_string(), Color::new(0, 0, 0)),
        ("lt1".to_string(), Color::new(255, 255, 255)),
        ("dk2".to_string(), Color::new(68, 84, 106)),
        ("lt2".to_string(), Color::new(231, 230, 230)),
        ("accent1".to_string(), Color::new(68, 114, 196)),
    ])
}

fn cap_join_theme() -> Vec<crate::parser::pptx::ThemeLineStyle> {
    crate::parser::pptx::parse_theme_line_styles(
        r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:fmtScheme><a:lnStyleLst><a:ln cap="rnd"><a:miter/></a:ln></a:lnStyleLst></a:fmtScheme></a:theme>"#,
    )
}

fn drawing_with_fill(color_markup: &str) -> String {
    format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:spPr>
        <a:solidFill>{color_markup}</a:solidFill>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p><a:r><a:t>hello</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#
    )
}

fn fill_of(color_markup: &str, theme: &HashMap<String, Color>) -> Option<Color> {
    let boxes = parse_drawing_text_boxes(
        &drawing_with_fill(color_markup),
        theme,
        &ThemeFontScheme::default(),
    );
    assert_eq!(boxes.len(), 1, "fixture should yield one text box");
    boxes[0].fill
}

#[test]
fn scheme_accent_fill_resolves_against_workbook_theme() {
    let fill = fill_of(r#"<a:schemeClr val="accent1"/>"#, &accent_theme());
    assert_eq!(fill, Some(Color::new(68, 114, 196)));
}

#[test]
fn scheme_fill_applies_lum_transforms() {
    // "accent1, lighter 60%" as Excel emits it: lumMod 40% + lumOff 60%.
    let fill = fill_of(
        r#"<a:schemeClr val="accent1"><a:lumMod val="40000"/><a:lumOff val="60000"/></a:schemeClr>"#,
        &accent_theme(),
    );
    // Matches the pptx transform math (tint/shade in RGB, lum in HSL).
    let expected = crate::parser::drawingml::apply_color_transforms(
        Color::new(68, 114, 196),
        &[
            crate::parser::drawingml::ColorTransform::LumMod(0.4),
            crate::parser::drawingml::ColorTransform::LumOff(0.6),
        ],
    );
    assert_eq!(fill, Some(expected));
}

#[test]
fn srgb_fill_with_shade_still_darkens() {
    // The old hand-rolled Empty(<a:shade>) path must keep working through the
    // shared parser.
    let fill = fill_of(
        r#"<a:srgbClr val="C86432"><a:shade val="50000"/></a:srgbClr>"#,
        &accent_theme(),
    );
    // 50% of #C86432's light, re-encoded (issue #667). Halving the sRGB bytes
    // would give (100, 50, 25).
    assert_eq!(fill, Some(Color::new(146, 71, 34)));
}

#[test]
fn background_scheme_names_use_spreadsheet_aliases() {
    // xlsx has no clrMap part; bg1/tx1 must map onto lt1/dk1.
    let theme = accent_theme();
    assert_eq!(
        fill_of(r#"<a:schemeClr val="bg1"/>"#, &theme),
        Some(Color::new(255, 255, 255))
    );
    assert_eq!(
        fill_of(r#"<a:schemeClr val="tx1"/>"#, &theme),
        Some(Color::new(0, 0, 0))
    );
    assert_eq!(
        fill_of(r#"<a:schemeClr val="bg2"/>"#, &theme),
        Some(Color::new(231, 230, 230))
    );
}

#[test]
fn scheme_fill_falls_back_to_light_dark_without_theme() {
    // Workbooks without a theme part keep the historical fallback.
    let empty = HashMap::new();
    assert_eq!(
        fill_of(r#"<a:schemeClr val="bg1"/>"#, &empty),
        Some(Color::new(255, 255, 255))
    );
    assert_eq!(
        fill_of(r#"<a:schemeClr val="tx1"/>"#, &empty),
        Some(Color::new(0, 0, 0))
    );
    assert_eq!(fill_of(r#"<a:schemeClr val="accent1"/>"#, &empty), None);
}

#[test]
fn theme_color_scheme_parses_from_theme_xml() {
    let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>"#;
    let colors = crate::parser::drawingml::parse_theme_color_scheme(theme_xml);
    assert_eq!(colors.get("dk1"), Some(&Color::new(0, 0, 0)));
    assert_eq!(colors.get("lt1"), Some(&Color::new(255, 255, 255)));
    assert_eq!(colors.get("dk2"), Some(&Color::new(0x44, 0x54, 0x6A)));
    assert_eq!(colors.get("accent1"), Some(&Color::new(0x44, 0x72, 0xC4)));
    assert_eq!(colors.get("hlink"), Some(&Color::new(0x05, 0x63, 0xC1)));
    assert_eq!(colors.get("accent2"), None);
}

fn office_theme_fonts() -> ThemeFontScheme {
    // The font scheme an Excel-saved workbook ships in theme1.xml.
    ThemeFontScheme {
        major_latin: Some("Calibri Light".to_string()),
        minor_latin: Some("Calibri".to_string()),
    }
}

fn drawing_with_run_properties(run_properties: &str) -> String {
    format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp>
      <xdr:txBody>
        <a:bodyPr/>
        <a:p><a:r>{run_properties}<a:t>accent1</a:t></a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#
    )
}

fn font_of(run_properties: &str, fonts: &ThemeFontScheme) -> Option<String> {
    let boxes = parse_drawing_text_boxes(
        &drawing_with_run_properties(run_properties),
        &HashMap::new(),
        fonts,
    );
    assert_eq!(boxes.len(), 1, "fixture should yield one text box");
    boxes[0].paragraphs[0].runs[0].style.font_family.clone()
}

#[test]
fn drawing_run_without_a_typeface_inherits_the_theme_minor_font() {
    // Excel writes shape labels as `<a:rPr lang="en-US" sz="1100"/>` with no
    // typeface at all; DrawingML resolves that to the theme's minor Latin
    // font, not to the renderer's serif default (issue #461).
    assert_eq!(
        font_of(r#"<a:rPr lang="en-US" sz="1100"/>"#, &office_theme_fonts()),
        Some("Calibri".to_string())
    );
}

#[test]
fn drawing_run_without_run_properties_inherits_the_theme_minor_font() {
    assert_eq!(
        font_of("", &office_theme_fonts()),
        Some("Calibri".to_string())
    );
}

#[test]
fn drawing_run_uses_its_explicit_latin_typeface() {
    assert_eq!(
        font_of(
            r#"<a:rPr lang="en-US" sz="1100"><a:latin typeface="Georgia"/></a:rPr>"#,
            &office_theme_fonts()
        ),
        Some("Georgia".to_string())
    );
}

#[test]
fn drawing_run_resolves_theme_typeface_placeholders() {
    // `+mj-lt` and `+mn-lt` name the theme's major and minor Latin fonts.
    assert_eq!(
        font_of(
            r#"<a:rPr lang="en-US"><a:latin typeface="+mj-lt"/></a:rPr>"#,
            &office_theme_fonts()
        ),
        Some("Calibri Light".to_string())
    );
    assert_eq!(
        font_of(
            r#"<a:rPr lang="en-US"><a:latin typeface="+mn-lt"/></a:rPr>"#,
            &office_theme_fonts()
        ),
        Some("Calibri".to_string())
    );
}

#[test]
fn drawing_text_box_outline_keeps_explicit_drawingml_cap_and_join() {
    let xml: String = drawing_with_fill(r#"<a:srgbClr val="FF0000"/>"#).replace(
        "</xdr:spPr>",
        r#"<a:ln cap="sq"><a:bevel/><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr><xdr:style><a:lnRef idx="1"/></xdr:style>"#,
    );
    let boxes = parse_drawing_text_boxes_with_theme_lines(
        &xml,
        &accent_theme(),
        &ThemeFontScheme::default(),
        &cap_join_theme(),
    );
    let border = boxes[0].border.as_ref().expect("text box has an outline");

    assert_eq!(border.cap, LineCap::Square);
    assert_eq!(border.join, LineJoin::Bevel);
}

#[test]
fn drawing_text_box_outline_inherits_cap_and_join_from_its_theme_line_reference() {
    let xml: String = drawing_with_fill(r#"<a:srgbClr val="FF0000"/>"#).replace(
        "</xdr:spPr>",
        r#"<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr><xdr:style><a:lnRef idx="1"/></xdr:style>"#,
    );
    let boxes = parse_drawing_text_boxes_with_theme_lines(
        &xml,
        &accent_theme(),
        &ThemeFontScheme::default(),
        &cap_join_theme(),
    );
    let border = boxes[0].border.as_ref().expect("text box has an outline");

    assert_eq!(border.cap, LineCap::Round);
    assert_eq!(border.join, LineJoin::Miter);
}

#[test]
fn drawing_run_font_stays_unset_without_a_theme_font_scheme() {
    // A workbook with no readable theme part leaves font selection to the
    // renderer's existing fallback rather than inventing a family.
    assert_eq!(
        font_of(r#"<a:rPr lang="en-US"/>"#, &ThemeFontScheme::default()),
        None
    );
}

#[test]
fn theme_font_scheme_parses_from_theme_xml() {
    let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>"#;
    let fonts = crate::parser::drawingml::parse_theme_font_scheme(theme_xml);
    assert_eq!(fonts.major_latin.as_deref(), Some("Calibri Light"));
    assert_eq!(fonts.minor_latin.as_deref(), Some("Calibri"));
}

#[test]
fn theme_font_scheme_ignores_empty_typefaces() {
    // Themes spell "inherit" as an empty typeface; it must not become a font.
    let theme_xml = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface=""/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>"#;
    let fonts = crate::parser::drawingml::parse_theme_font_scheme(theme_xml);
    assert_eq!(fonts, ThemeFontScheme::default());
}

/// `(font_size, bold, italic)` of the single run the fixture produces.
fn run_emphasis(run_properties: &str) -> (Option<f64>, Option<bool>, Option<bool>) {
    let boxes = parse_drawing_text_boxes(
        &drawing_with_run_properties(run_properties),
        &HashMap::new(),
        &ThemeFontScheme::default(),
    );
    assert_eq!(boxes.len(), 1, "fixture should yield one text box");
    let style = &boxes[0].paragraphs[0].runs[0].style;
    (style.font_size, style.bold, style.italic)
}

#[test]
fn self_closing_run_properties_still_carry_size_and_emphasis() {
    // Excel writes shape-label run properties as a self-closing element,
    // which quick-xml reports as `Event::Empty`; reading them only on
    // `Event::Start` dropped every attribute (issue #466).
    assert_eq!(
        run_emphasis(r#"<a:rPr lang="en-US" sz="1400" b="1" i="1"/>"#),
        (Some(14.0), Some(true), Some(true))
    );
}

#[test]
fn run_properties_with_children_keep_working() {
    // Triangulation: the same attributes on an element that has children
    // must resolve identically, so the fix cannot special-case one spelling.
    assert_eq!(
        run_emphasis(
            r#"<a:rPr lang="en-US" sz="1400" b="1" i="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr>"#
        ),
        (Some(14.0), Some(true), Some(true))
    );
}

#[test]
fn run_properties_omit_emphasis_when_disabled() {
    // `b="0"`/`i="0"` mean "not bold"/"not italic"; they must not set the flag.
    assert_eq!(
        run_emphasis(r#"<a:rPr lang="en-US" sz="900" b="0" i="0"/>"#),
        (Some(9.0), None, None)
    );
}

/// A worksheet drawing holding one picture anchor whose blip is spelled with
/// the given markup.
fn drawing_with_blip(blip_markup: &str) -> String {
    format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>11880</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>360720</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>963720</xdr:colOff><xdr:row>11</xdr:row><xdr:rowOff>5760</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 10" descr="Three wrapped gifts"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill>{blip_markup}<a:stretch/></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="156600" y="5970960"/><a:ext cx="3268440" cy="2178720"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#
    )
}

#[test]
fn self_closing_blip_yields_its_embed_relationship() {
    let anchors = parse_drawing_image_anchors(&drawing_with_blip(r#"<a:blip r:embed="rId2"/>"#));
    assert_eq!(anchors.len(), 1, "the anchor holds one picture");
    assert_eq!(anchors[0].1.rid, "rId2");
    assert_eq!(anchors[0].0.from_row, 5);
    assert_eq!(anchors[0].0.from_col, 1);
}

#[test]
fn blip_without_a_transparency_effect_has_no_alpha() {
    let anchors = parse_drawing_image_anchors(&drawing_with_blip(r#"<a:blip r:embed="rId2"/>"#));
    assert_eq!(
        anchors[0].1.alpha, None,
        "a picture with no <a:alphaModFix> prints at full strength"
    );
}

#[test]
fn blip_with_effect_children_still_yields_its_embed_relationship() {
    // A picture carrying an alpha or recolour effect writes `a:blip` as a start
    // element with children; reading `r:embed` only off `Event::Empty` dropped
    // the whole picture (issue #1066). Markup is the reported workbook's.
    let anchors = parse_drawing_image_anchors(&drawing_with_blip(
        r#"<a:blip r:embed="rId2"><a:alphaModFix amt="70000"/><a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a14:imgLayer r:embed="rId3"><a14:imgEffect><a14:brightnessContrast colorTemp="5513" contrast="40000" sat="0"/></a14:imgEffect></a14:imgLayer></a14:imgProps></a:ext></a:extLst></a:blip>"#,
    ));
    assert_eq!(anchors.len(), 1, "the anchor still holds one picture");
    // rId3 is the a14 high-definition layer, not the picture Excel draws.
    assert_eq!(anchors[0].1.rid, "rId2");
    assert_eq!(anchors[0].0.from_row, 5);
    assert_eq!(anchors[0].0.from_col, 1);
}

/// `<a:alphaModFix amt="70000"/>` makes Excel draw the picture at 70% of its
/// strength, compositing it onto whatever the worksheet shows beneath. The
/// amount rides the same blip as `r:embed`, so it has to be picked up while
/// that element is open (issue #1103).
#[test]
fn blip_alpha_mod_fix_is_read_as_a_transparency_factor() {
    let anchors = parse_drawing_image_anchors(&drawing_with_blip(
        r#"<a:blip r:embed="rId2"><a:alphaModFix amt="70000"/></a:blip>"#,
    ));
    assert_eq!(anchors.len(), 1, "the anchor holds one picture");
    assert_eq!(anchors[0].1.rid, "rId2");
    assert_eq!(
        anchors[0].1.alpha,
        Some(0.7),
        "amt is a percentage in thousandths of a percent"
    );
}

/// `<a:alphaModFix>` holds nothing but attributes, so a writer is free to
/// spell it with an explicit end tag. That reaches the reader as a start/end
/// pair rather than one empty element, exactly the spelling difference that
/// dropped effect-bearing pictures in #1066.
#[test]
fn blip_alpha_mod_fix_written_with_an_explicit_end_tag_is_read() {
    let anchors = parse_drawing_image_anchors(&drawing_with_blip(
        r#"<a:blip r:embed="rId2"><a:alphaModFix amt="45000"></a:alphaModFix></a:blip>"#,
    ));
    assert_eq!(anchors.len(), 1, "the anchor holds one picture");
    assert_eq!(anchors[0].1.alpha, Some(0.45));
}

/// The `a14:imgProps` extension Excel nests inside an effect-bearing blip
/// carries its own effect markup for the high-definition layer. Only the
/// blip's own `<a:alphaModFix>` describes the picture Excel draws, the same
/// rule `r:embed` already follows (issue #1066).
#[test]
fn alpha_mod_fix_inside_the_a14_extension_is_not_the_pictures_alpha() {
    let anchors = parse_drawing_image_anchors(&drawing_with_blip(
        r#"<a:blip r:embed="rId2"><a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><a14:imgLayer r:embed="rId3"><a14:imgEffect><a:alphaModFix amt="20000"/></a14:imgEffect></a14:imgLayer></a14:imgProps></a:ext></a:extLst></a:blip>"#,
    ));
    assert_eq!(anchors.len(), 1, "the anchor holds one picture");
    assert_eq!(
        anchors[0].1.alpha, None,
        "the extension layer's amount must not become the picture's alpha"
    );
}

/// One worksheet `_rels` part declaring a grid drawing, a header/footer
/// drawing and a legacy VML drawing — the three relationship types whose URIs
/// a `contains("drawing")` test cannot tell apart.
const SHEET_RELS_WITH_THREE_DRAWING_KINDS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawingHF" Target="../drawings/drawing2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
</Relationships>"#;

/// A worksheet body naming the given elements after its `<sheetData>`.
fn worksheet_with(elements: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
  {elements}
</worksheet>"#
    )
}

#[test]
fn a_sheet_drawing_element_resolves_its_relationship_target() {
    let targets = sheet_drawing_targets(
        &worksheet_with(r#"<drawing r:id="rId1"/>"#),
        SHEET_RELS_WITH_THREE_DRAWING_KINDS,
    );
    assert_eq!(targets, vec!["../drawings/drawing1.xml".to_string()]);
}

/// The defect of issue #1158: the relationship survives, the drawing part and
/// its media survive, and only the `<drawing>` element is gone. Excel prints
/// the grid alone, so nothing must resolve.
#[test]
fn a_drawing_relationship_the_sheet_never_names_resolves_to_nothing() {
    let targets = sheet_drawing_targets(
        &worksheet_with(r#"<pageSetup paperSize="9" orientation="portrait"/>"#),
        SHEET_RELS_WITH_THREE_DRAWING_KINDS,
    );
    assert!(
        targets.is_empty(),
        "an unreferenced relationship attaches nothing, got {targets:?}"
    );
}

/// `drawingHF` and `vmlDrawing` both contain `drawing` in their relationship
/// type. Neither is a grid drawing, and each is named by an element of its own.
#[test]
fn header_footer_and_legacy_drawings_are_not_grid_drawings() {
    let targets = sheet_drawing_targets(
        &worksheet_with(r#"<drawingHF r:id="rId2"/><legacyDrawing r:id="rId3"/>"#),
        SHEET_RELS_WITH_THREE_DRAWING_KINDS,
    );
    assert!(
        targets.is_empty(),
        "only `<drawing>` puts a part on the grid, got {targets:?}"
    );
}

/// A sheet may name more than one drawing, and the body's order is the paint
/// order — not whatever order the `_rels` part happens to list.
#[test]
fn multiple_drawing_elements_resolve_in_body_order() {
    let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rIdA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
      <Relationship Id="rIdB" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/>
    </Relationships>"#;
    let targets = sheet_drawing_targets(
        &worksheet_with(r#"<drawing r:id="rIdB"/><drawing r:id="rIdA"/>"#),
        rels,
    );
    assert_eq!(
        targets,
        vec![
            "../drawings/drawing2.xml".to_string(),
            "../drawings/drawing1.xml".to_string()
        ]
    );
}

// ── Line shapes (issue #1566) ────────────────────────────────────────────

/// A worksheet drawing holding one line-geometry shape, as Excel writes the
/// "Chart border" connectors of the budget workbook on issue #1566.
fn drawing_with_line_shape(element: &str, shape_properties: &str, style: &str) -> String {
    format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>482419</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>28322</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>482419</xdr:colOff><xdr:row>14</xdr:row><xdr:rowOff>47387</xdr:rowOff></xdr:to>
    <xdr:{element} macro="">
      <xdr:nvCxnSpPr><xdr:cNvPr id="19" name="Chart border 1" descr="Chart border"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>
      <xdr:spPr>{shape_properties}</xdr:spPr>
      {style}
    </xdr:{element}>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#
    )
}

const LINE_XFRM: &str = r#"<a:xfrm><a:off x="4778194" y="1780922"/><a:ext cx="0" cy="2524140"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>"#;

const ACCENT_LINE_REFERENCE: &str = r#"<xdr:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></xdr:style>"#;

#[test]
fn connector_line_shape_yields_its_anchor_stroke_and_colour() {
    let xml = drawing_with_line_shape(
        "cxnSp",
        &format!(
            r#"{LINE_XFRM}<a:ln w="12700"><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr></a:solidFill></a:ln>"#
        ),
        ACCENT_LINE_REFERENCE,
    );
    let lines = parse_drawing_line_shapes(&xml, &accent_theme());
    assert_eq!(lines.len(), 1, "the anchor holds one line shape");
    let line = &lines[0];
    assert_eq!(line.geometry.from_col, 3);
    assert_eq!(line.geometry.from_col_off_emu, 482_419);
    assert_eq!(line.geometry.from_row, 3);
    assert_eq!(line.geometry.from_row_off_emu, 28_322);
    assert_eq!(line.geometry.to, Some((3, 482_419, 14, 47_387)));
    // `a:ln w="12700"` is one point; "bg1, darker 15%" is #D9D9D9.
    assert!(
        (line.stroke.width - 1.0).abs() < 1e-9,
        "{}",
        line.stroke.width
    );
    assert_eq!(line.stroke.color, Color::new(217, 217, 217));
    assert!(!line.flip_h && !line.flip_v);
}

#[test]
fn worksheet_line_shape_keeps_explicit_drawingml_cap_and_join() {
    let xml = drawing_with_line_shape(
        "cxnSp",
        &format!(
            r#"{LINE_XFRM}<a:ln w="12700" cap="sq"><a:bevel/><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>"#
        ),
        ACCENT_LINE_REFERENCE,
    );
    let lines =
        parse_drawing_line_shapes_with_theme_lines(&xml, &accent_theme(), &cap_join_theme());

    assert_eq!(lines[0].stroke.cap, LineCap::Square);
    assert_eq!(lines[0].stroke.join, LineJoin::Bevel);
}

#[test]
fn worksheet_line_shape_inherits_cap_and_join_from_its_theme_line_reference() {
    let xml = drawing_with_line_shape(
        "cxnSp",
        &format!(
            r#"{LINE_XFRM}<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>"#
        ),
        ACCENT_LINE_REFERENCE,
    );
    let lines =
        parse_drawing_line_shapes_with_theme_lines(&xml, &accent_theme(), &cap_join_theme());

    assert_eq!(lines[0].stroke.cap, LineCap::Round);
    assert_eq!(lines[0].stroke.join, LineJoin::Miter);
}

#[test]
fn flipped_connector_records_its_mirrored_direction() {
    let xml = drawing_with_line_shape(
        "cxnSp",
        r#"<a:xfrm flipV="1"><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>"#,
        "",
    );
    let lines = parse_drawing_line_shapes(&xml, &accent_theme());
    assert_eq!(lines.len(), 1, "a straight connector is a line shape too");
    assert!(
        lines[0].flip_v,
        "flipV runs the line bottom-left to top-right"
    );
    assert!(!lines[0].flip_h);
    assert!((lines[0].stroke.width - 1.5).abs() < 1e-9);
    assert_eq!(lines[0].stroke.color, Color::new(255, 0, 0));
}

#[test]
fn connector_without_a_line_fill_takes_the_style_line_reference() {
    // Excel's default connector: an empty `<a:ln/>` whose colour comes from
    // `<xdr:style><a:lnRef>`, at DrawingML's 0.75pt default width.
    let xml = drawing_with_line_shape(
        "cxnSp",
        &format!("{LINE_XFRM}<a:ln/>"),
        ACCENT_LINE_REFERENCE,
    );
    let lines = parse_drawing_line_shapes(&xml, &accent_theme());
    assert_eq!(lines.len(), 1);
    assert_eq!(lines[0].stroke.color, Color::new(68, 114, 196));
    assert!((lines[0].stroke.width - 0.75).abs() < 1e-9);
}

#[test]
fn connector_whose_line_is_no_fill_is_not_drawn() {
    let xml = drawing_with_line_shape(
        "cxnSp",
        &format!("{LINE_XFRM}<a:ln w=\"12700\"><a:noFill/></a:ln>"),
        ACCENT_LINE_REFERENCE,
    );
    assert!(parse_drawing_line_shapes(&xml, &accent_theme()).is_empty());
}

#[test]
fn hidden_line_shape_is_not_drawn() {
    let xml = drawing_with_line_shape(
        "cxnSp",
        &format!(
            r#"{LINE_XFRM}<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>"#
        ),
        "",
    )
    .replace(r#"name="Chart border 1""#, r#"name="Chart border 1" hidden="1""#);
    assert!(parse_drawing_line_shapes(&xml, &accent_theme()).is_empty());
}

#[test]
fn a_plain_shape_with_line_geometry_is_a_line_shape() {
    let xml = drawing_with_line_shape(
        "sp",
        &format!(
            r#"{LINE_XFRM}<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>"#
        ),
        "",
    );
    let lines = parse_drawing_line_shapes(&xml, &accent_theme());
    assert_eq!(lines.len(), 1, "`xdr:sp` with `prst=\"line\"` is a line");
    assert!((lines[0].stroke.width - 0.75).abs() < 1e-9);
}

#[test]
fn a_text_box_rectangle_is_not_a_line_shape() {
    let xml = drawing_with_fill(r#"<a:srgbClr val="FF0000"/>"#);
    assert!(parse_drawing_line_shapes(&xml, &accent_theme()).is_empty());
}
