use std::collections::HashMap;

use super::*;
use crate::ir::Color;

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
    let boxes = parse_drawing_text_boxes(&drawing_with_fill(color_markup), theme);
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
    assert_eq!(fill, Some(Color::new(100, 50, 25)));
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
