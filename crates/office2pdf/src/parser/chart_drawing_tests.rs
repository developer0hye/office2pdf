use super::*;

use std::collections::HashMap;
use std::sync::LazyLock;

use crate::ir::{ChartUserShapeExtent, Color, LineCap, LineJoin};

static NO_ALIASES: LazyLock<HashMap<String, String>> = LazyLock::new(HashMap::new);

/// The palette `Monthly college budget1.xlsx` resolves its caption against:
/// `accent1` is `#67BCD1`, which the run's `<a:lumMod val="50000"/>` halves the
/// luminance of.
static WORKBOOK_THEME: LazyLock<HashMap<String, Color>> = LazyLock::new(|| {
    HashMap::from([
        ("accent1".to_string(), Color::new(0x67, 0xBC, 0xD1)),
        ("dk1".to_string(), Color::new(0, 0, 0)),
    ])
});

fn workbook_scheme() -> SchemeColors<'static> {
    SchemeColors {
        colors: &WORKBOOK_THEME,
        aliases: &NO_ALIASES,
    }
}

fn cambria_theme() -> ThemeFontScheme {
    ThemeFontScheme {
        major_latin: Some("Cambria".to_string()),
        minor_latin: Some("Trebuchet MS".to_string()),
    }
}

fn cap_join_theme() -> Vec<crate::parser::pptx::ThemeLineStyle> {
    crate::parser::pptx::parse_theme_line_styles(
        r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:fmtScheme><a:lnStyleLst><a:ln cap="rnd"><a:miter/></a:ln></a:lnStyleLst></a:fmtScheme></a:theme>"#,
    )
}

/// `xl/drawings/drawing2.xml` of `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`,
/// which `xl/charts/chart2.xml` names through `<c:userShapes r:id="rId3"/>`, cut
/// down to the elements this parser reads. The caption it holds is the
/// `CASH FLOW` heading Excel prints left of that chart's plot (issue #1186).
const CASH_FLOW_CAPTION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <cdr:relSizeAnchor xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing">
    <cdr:from><cdr:x>0</cdr:x><cdr:y>0.17913</cdr:y></cdr:from>
    <cdr:to><cdr:x>0.11685</cdr:x><cdr:y>0.50958</cdr:y></cdr:to>
    <cdr:sp macro="" textlink="">
      <cdr:nvSpPr><cdr:cNvPr id="2" name="TextBox 3"/><cdr:cNvSpPr txBox="1"/></cdr:nvSpPr>
      <cdr:spPr>
        <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:off x="0" y="172324"/><a:ext cx="1230978" cy="317908"/></a:xfrm>
        <a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>
      </cdr:spPr>
      <cdr:txBody>
        <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" wrap="none" rtlCol="0" anchor="t"><a:spAutoFit/></a:bodyPr>
        <a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:lvl1pPr marL="0" indent="0"><a:defRPr sz="1100"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></a:lstStyle>
        <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:pPr algn="l"/>
          <a:r>
            <a:rPr lang="en-US" sz="1500" b="1"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill><a:latin typeface="+mj-lt"/></a:rPr>
            <a:t>CASH FLOW</a:t>
          </a:r>
        </a:p>
      </cdr:txBody>
    </cdr:sp>
  </cdr:relSizeAnchor>
</c:userShapes>"#;

#[test]
fn rel_size_anchor_states_both_corners_as_chart_area_fractions() {
    let shapes = parse_chart_user_shapes(CASH_FLOW_CAPTION, &workbook_scheme(), &cambria_theme());

    assert_eq!(shapes.len(), 1, "the part holds one anchored shape");
    let caption = &shapes[0];
    assert_eq!(caption.from, (0.0, 0.17913));
    assert_eq!(
        caption.extent,
        ChartUserShapeExtent::Corner {
            x: 0.11685,
            y: 0.50958
        }
    );
}

#[test]
fn shape_text_carries_its_run_size_weight_face_and_resolved_scheme_color() {
    let shapes = parse_chart_user_shapes(CASH_FLOW_CAPTION, &workbook_scheme(), &cambria_theme());
    let caption = &shapes[0];

    assert_eq!(caption.paragraphs.len(), 1);
    let runs = &caption.paragraphs[0].runs;
    assert_eq!(runs.len(), 1);
    assert_eq!(runs[0].text, "CASH FLOW");
    assert_eq!(runs[0].style.font_size, Some(15.0));
    assert_eq!(runs[0].style.bold, Some(true));
    // `+mj-lt` is the theme's major Latin face, as it is anywhere else in
    // DrawingML; the chart part names no theme of its own, so the package's
    // reaches this parser through its caller.
    assert_eq!(runs[0].style.font_family.as_deref(), Some("Cambria"));
    // `accent1` at half the luminance: Excel's own export of this caption
    // prints it in `#246778`.
    assert_eq!(runs[0].style.color, Some(Color::new(0x24, 0x67, 0x78)));
}

#[test]
fn chart_user_shape_outline_keeps_explicit_drawingml_cap_and_join() {
    let xml: String = CASH_FLOW_CAPTION.replace(
        "</cdr:spPr>",
        r#"<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" cap="sq"><a:bevel/><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></cdr:spPr><cdr:style><a:lnRef xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" idx="1"/></cdr:style>"#,
    );
    let shapes = parse_chart_user_shapes_with_theme_lines(
        &xml,
        &workbook_scheme(),
        &cambria_theme(),
        &cap_join_theme(),
    );
    let border = shapes[0]
        .border
        .as_ref()
        .expect("chart text box has an outline");

    assert_eq!(border.cap, LineCap::Square);
    assert_eq!(border.join, LineJoin::Bevel);
}

#[test]
fn chart_user_shape_outline_inherits_cap_and_join_from_its_theme_line_reference() {
    let xml: String = CASH_FLOW_CAPTION.replace(
        "</cdr:spPr>",
        r#"<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></cdr:spPr><cdr:style><a:lnRef xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" idx="1"/></cdr:style>"#,
    );
    let shapes = parse_chart_user_shapes_with_theme_lines(
        &xml,
        &workbook_scheme(),
        &cambria_theme(),
        &cap_join_theme(),
    );
    let border = shapes[0]
        .border
        .as_ref()
        .expect("chart text box has an outline");

    assert_eq!(border.cap, LineCap::Round);
    assert_eq!(border.join, LineJoin::Miter);
}

/// `<a:bodyPr>` states the insets in EMU when it states them at all; the
/// DrawingML defaults are 0.1in left and right and 0.05in top and bottom, and
/// a native Excel export seats this caption's pen exactly 7.2pt inside its box.
#[test]
fn body_insets_default_to_the_drawingml_pair_and_a_stated_value_wins() {
    let defaulted =
        parse_chart_user_shapes(CASH_FLOW_CAPTION, &workbook_scheme(), &cambria_theme());
    assert_eq!(defaulted[0].text_insets.left, 7.2);
    assert_eq!(defaulted[0].text_insets.top, 3.6);

    let stated: String = CASH_FLOW_CAPTION.replace(
        r#"wrap="none" rtlCol="0" anchor="t""#,
        r#"wrap="none" rtlCol="0" anchor="t" lIns="0" tIns="0""#,
    );
    let shapes = parse_chart_user_shapes(&stated, &workbook_scheme(), &cambria_theme());
    assert_eq!(shapes[0].text_insets.left, 0.0);
    assert_eq!(shapes[0].text_insets.top, 0.0);
}

/// `<cdr:absSizeAnchor>` gives the size in EMU rather than as a second corner,
/// so the shape keeps that size however large the chart is.
#[test]
fn abs_size_anchor_states_its_extent_in_points() {
    let xml: String = CASH_FLOW_CAPTION
        .replace("relSizeAnchor", "absSizeAnchor")
        .replace(
            "<cdr:to><cdr:x>0.11685</cdr:x><cdr:y>0.50958</cdr:y></cdr:to>",
            r#"<cdr:ext cx="1230978" cy="317908"/>"#,
        );

    let shapes = parse_chart_user_shapes(&xml, &workbook_scheme(), &cambria_theme());
    assert_eq!(shapes.len(), 1);
    let ChartUserShapeExtent::Size { width, height } = shapes[0].extent else {
        panic!("an absSizeAnchor states a size, not a corner");
    };
    assert!((width - 96.927).abs() < 0.001, "1230978 EMU in points");
    assert!((height - 25.032).abs() < 0.001, "317908 EMU in points");
}

/// A shape with no text at all still counts: it is drawn, and dropping it here
/// would lose whatever fill or outline it carries.
#[test]
fn a_shape_naming_no_text_still_reaches_the_ir_with_its_fill() {
    let xml: String = CASH_FLOW_CAPTION
        .replace("<a:t>CASH FLOW</a:t>", "<a:t></a:t>")
        .replace(
            r#"<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#,
            r#"<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="FF0000"/></a:solidFill>"#,
        );

    let shapes = parse_chart_user_shapes(&xml, &workbook_scheme(), &cambria_theme());
    assert_eq!(shapes.len(), 1);
    assert_eq!(shapes[0].fill, Some(Color::new(0xFF, 0, 0)));
}

/// The relationship the chart part names, so the loader reads the drawing the
/// chart actually points at rather than every `chartUserShapes` relationship
/// the package happens to declare (the reverse of issue #1158).
#[test]
fn the_chart_part_names_the_relationship_its_user_shapes_come_from() {
    let chart_xml: &str = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <c:chart><c:plotArea/></c:chart>
            <c:userShapes r:id="rId3"/>
        </c:chartSpace>"#;
    assert_eq!(user_shapes_rid(chart_xml).as_deref(), Some("rId3"));

    let without: &str = r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea/></c:chart>
        </c:chartSpace>"#;
    assert_eq!(user_shapes_rid(without), None);
}
