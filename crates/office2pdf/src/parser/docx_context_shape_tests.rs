use super::*;
use crate::ir::{ArrowHead, LineCap, LineJoin, ShapeKind};

/// A document.xml body wrapper around `inner` run/drawing markup.
fn body(inner: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<w:body><w:p><w:r>{inner}</w:r></w:p></w:body></w:document>"#
    )
}

/// A filled rectangle authored by LibreOffice (issue #176, "Shape 2").
const RECT_DRAWING: &str = r#"<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
<wp:anchor>
<wp:positionH relativeFrom="column"><wp:posOffset>2886710</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>141605</wp:posOffset></wp:positionV>
<wp:extent cx="1590675" cy="733425"/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp><wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="1590840" cy="733320"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
<a:solidFill><a:srgbClr val="729fcf"/></a:solidFill>
<a:ln w="0"><a:solidFill><a:srgbClr val="3465a4"/></a:solidFill></a:ln>
</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>
<mc:Fallback><w:pict><v:rect/></w:pict></mc:Fallback></mc:AlternateContent>"#;

/// A horizontal connector with a triangular arrowhead (issue #176, "Horizontal line 1").
const LINE_DRAWING: &str = r#"<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
<wp:anchor>
<wp:positionH relativeFrom="column"><wp:posOffset>1957070</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>-74295</wp:posOffset></wp:positionV>
<wp:extent cx="929640" cy="0"/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp><wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="929520" cy="0"/></a:xfrm>
<a:prstGeom prst="line"><a:avLst/></a:prstGeom>
<a:ln w="0"><a:solidFill><a:srgbClr val="3465a4"/></a:solidFill>
<a:tailEnd len="med" type="triangle" w="med"/></a:ln>
</wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>
<mc:Fallback><w:pict><v:line/></w:pict></mc:Fallback></mc:AlternateContent>"#;

/// A text-box shape (`wps:txbx`) — handled by docx-rs, must be ignored here.
const TEXTBOX_DRAWING: &str = r#"<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
<wp:anchor>
<wp:positionH relativeFrom="column"><wp:posOffset>2985770</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>-25400</wp:posOffset></wp:positionV>
<wp:extent cx="1390650" cy="485775"/>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp><wps:spPr>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="0"><a:noFill/></a:ln>
</wps:spPr>
<wps:txbx><w:txbxContent><w:p><w:r><w:t>Very important text inside a box</w:t></w:r></w:p></w:txbxContent></wps:txbx>
</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent>"#;

/// An inline picture — handled by docx-rs, must be ignored here.
const PIC_DRAWING: &str = r#"<w:drawing><wp:anchor>
<wp:positionH relativeFrom="column"><wp:posOffset>60325</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="paragraph"><wp:posOffset>635</wp:posOffset></wp:positionV>
<wp:extent cx="3432175" cy="2574290"/>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic><pic:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
</a:graphicData></a:graphic></wp:anchor></w:drawing>"#;

fn approx(actual: f64, expected: f64) {
    assert!(
        (actual - expected).abs() < 0.05,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn scans_filled_rectangle_geometry_position_and_colors() {
    let shapes = scan_drawing_shapes(&body(RECT_DRAWING));
    assert_eq!(shapes.len(), 1, "expected one rectangle shape");

    let shape = &shapes[0];
    assert!(matches!(shape.shape.kind, ShapeKind::Rectangle));
    approx(shape.offset_x, 227.30); // 2886710 EMU
    approx(shape.offset_y, 11.15); // 141605 EMU
    approx(shape.width, 125.25); // 1590675 EMU
    approx(shape.height, 57.75); // 733425 EMU

    let fill = shape.shape.fill.expect("rectangle should have a fill");
    assert_eq!((fill.r, fill.g, fill.b), (0x72, 0x9f, 0xcf));

    let stroke = shape
        .shape
        .stroke
        .as_ref()
        .expect("rectangle has an outline");
    assert_eq!(
        (stroke.color.r, stroke.color.g, stroke.color.b),
        (0x34, 0x65, 0xa4)
    );
    assert!(stroke.width > 0.0, "w=0 must map to a visible hairline");
}

#[test]
fn scans_line_with_tail_arrowhead() {
    let shapes = scan_drawing_shapes(&body(LINE_DRAWING));
    assert_eq!(shapes.len(), 1, "expected one line shape");

    let shape = &shapes[0];
    match shape.shape.kind {
        ShapeKind::Line {
            head_end, tail_end, ..
        } => {
            assert_eq!(tail_end, ArrowHead::Triangle, "tailEnd triangle → arrow");
            assert_eq!(head_end, ArrowHead::None);
        }
        ref other => panic!("expected a line, got {other:?}"),
    }
    assert!(shape.shape.fill.is_none(), "a line has no fill");
    let stroke = shape.shape.stroke.as_ref().expect("line needs a stroke");
    assert_eq!(
        (stroke.color.r, stroke.color.g, stroke.color.b),
        (0x34, 0x65, 0xa4)
    );
}

#[test]
fn shape_outline_keeps_explicit_drawingml_cap_and_join() {
    let drawing: String = RECT_DRAWING
        .replace("<a:ln w=\"0\">", "<a:ln w=\"0\" cap=\"sq\"><a:bevel/>")
        .replace(
            "</wps:spPr>",
            "</wps:spPr><wps:style><a:lnRef idx=\"1\"/></wps:style>",
        );
    let xml: String = body(&drawing);
    let theme: &str = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:fmtScheme><a:lnStyleLst><a:ln cap="rnd"><a:miter/></a:ln></a:lnStyleLst></a:fmtScheme></a:theme>"#;
    let context: DrawingShapeContext =
        DrawingShapeContext::from_xml_with_theme(Some(&xml), Some(theme));
    let shape: FloatingShape = context.consume_next().expect("styled shape is kept");
    let stroke = shape
        .shape
        .stroke
        .as_ref()
        .expect("rectangle has an outline");

    assert_eq!(stroke.cap, LineCap::Square);
    assert_eq!(stroke.join, LineJoin::Bevel);
}

#[test]
fn shape_outline_inherits_cap_and_join_from_its_theme_line_reference() {
    let drawing: String = RECT_DRAWING.replace(
        "</wps:spPr>",
        "</wps:spPr><wps:style><a:lnRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:lnRef></wps:style>",
    );
    let xml: String = body(&drawing);
    let theme: &str = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:fmtScheme><a:lnStyleLst><a:ln cap="rnd"><a:miter/></a:ln></a:lnStyleLst></a:fmtScheme></a:theme>"#;
    let context: DrawingShapeContext =
        DrawingShapeContext::from_xml_with_theme(Some(&xml), Some(theme));
    let shape: FloatingShape = context.consume_next().expect("theme-styled shape is kept");
    let stroke = shape.shape.stroke.as_ref().expect("shape has an outline");

    assert_eq!(stroke.cap, LineCap::Round);
    assert_eq!(stroke.join, LineJoin::Miter);
}

#[test]
fn text_box_frame_inherits_cap_and_join_from_its_theme_line_reference() {
    let drawing: String = TEXTBOX_DRAWING
        .replace(
            "<a:ln w=\"0\"><a:noFill/></a:ln>",
            "<a:ln w=\"0\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln>",
        )
        .replace(
            "</wps:spPr>",
            "</wps:spPr><wps:style><a:lnRef idx=\"1\"/></wps:style>",
        );
    let xml: String = body(&drawing);
    let theme: &str = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:fmtScheme><a:lnStyleLst><a:ln cap="rnd"><a:miter/></a:ln></a:lnStyleLst></a:fmtScheme></a:theme>"#;
    let context: super::super::drawing::DrawingTextBoxContext =
        super::super::drawing::DrawingTextBoxContext::from_xml_with_theme(Some(&xml), Some(theme));
    let text_box = context.consume_next();
    let stroke = text_box.stroke.as_ref().expect("text box has an outline");

    assert_eq!(stroke.cap, LineCap::Round);
    assert_eq!(stroke.join, LineJoin::Miter);
}

#[test]
fn ignores_text_box_and_picture_drawings() {
    // Text boxes and pictures are handled by docx-rs; this side-channel must
    // not double-emit them.
    assert!(scan_drawing_shapes(&body(TEXTBOX_DRAWING)).is_empty());
    assert!(scan_drawing_shapes(&body(PIC_DRAWING)).is_empty());
}

#[test]
fn scans_multiple_shapes_in_document_order() {
    let combined = format!("{RECT_DRAWING}{TEXTBOX_DRAWING}{LINE_DRAWING}");
    let shapes = scan_drawing_shapes(&body(&combined));
    // Only the two geometry-only shapes survive, in order: rect then line.
    assert_eq!(shapes.len(), 2);
    assert!(matches!(shapes[0].shape.kind, ShapeKind::Rectangle));
    assert!(matches!(shapes[1].shape.kind, ShapeKind::Line { .. }));
}

#[test]
fn consume_next_yields_shapes_then_none() {
    let ctx = DrawingShapeContext::from_xml(Some(&body(RECT_DRAWING)));
    assert!(ctx.consume_next().is_some());
    assert!(
        ctx.consume_next().is_none(),
        "cursor past the end yields None"
    );
}

#[test]
fn empty_when_no_drawings() {
    assert!(scan_drawing_shapes(&body("<w:t>plain text</w:t>")).is_empty());
}

#[test]
fn wpg_explicit_colors_outrank_zero_index_style_references() {
    let drawing = r#"<w:drawing><wp:anchor>
<wp:positionH><wp:posOffset>0</wp:posOffset></wp:positionH>
<wp:positionV><wp:posOffset>0</wp:posOffset></wp:positionV>
<a:graphic><a:graphicData><wpg:wgp><wps:wsp>
<wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/></a:xfrm>
<a:prstGeom prst="rect"/>
<a:solidFill><a:schemeClr val="accent2"/></a:solidFill>
<a:ln w="12700"><a:solidFill><a:schemeClr val="accent3"/></a:solidFill></a:ln>
</wps:spPr>
<wps:style>
<a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>
<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>
</wps:style>
</wps:wsp></wpg:wgp></a:graphicData></a:graphic>
</wp:anchor></w:drawing>"#;
    let theme = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<a:accent1><a:srgbClr val="112233"/></a:accent1>
<a:accent2><a:srgbClr val="445566"/></a:accent2>
<a:accent3><a:srgbClr val="778899"/></a:accent3>
</a:theme>"#;

    let records = scan_wpg_drawings(&body(drawing), Some(theme));
    let shape = records[0].as_ref().expect("WPG drawing").children[0]
        .shape
        .as_ref()
        .expect("WPG shape");

    assert_eq!(shape.shape.fill, Some(Color::new(0x44, 0x55, 0x66)));
    assert_eq!(
        shape.shape.stroke.as_ref().map(|stroke| stroke.color),
        Some(Color::new(0x77, 0x88, 0x99))
    );
}

#[test]
fn wpg_theme_aliases_fill_alpha_and_transforms_use_shared_color_resolution() {
    let drawing = r#"<w:drawing><wp:anchor>
<wp:positionH><wp:posOffset>0</wp:posOffset></wp:positionH>
<wp:positionV><wp:posOffset>0</wp:posOffset></wp:positionV>
<a:graphic><a:graphicData><wpg:wgp>
<wps:wsp><wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/></a:xfrm>
<a:prstGeom prst="rect"/>
<a:solidFill><a:schemeClr val="tx1"><a:alpha val="16000"/></a:schemeClr></a:solidFill>
<a:ln><a:noFill/></a:ln>
</wps:spPr><wps:style><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef></wps:style></wps:wsp>
<wps:wsp><wps:spPr>
<a:xfrm><a:off x="127000" y="0"/><a:ext cx="127000" cy="127000"/></a:xfrm>
<a:prstGeom prst="rect"/>
<a:solidFill><a:schemeClr val="bg1"><a:shade val="50000"/></a:schemeClr></a:solidFill>
<a:ln><a:noFill/></a:ln>
</wps:spPr><wps:style><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef></wps:style></wps:wsp>
</wpg:wgp></a:graphicData></a:graphic>
</wp:anchor></w:drawing>"#;
    let theme = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<a:themeElements><a:clrScheme name="Test">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:accent1><a:srgbClr val="B2C1D8"/></a:accent1>
</a:clrScheme></a:themeElements>
</a:theme>"#;

    let records = scan_wpg_drawings(&body(drawing), Some(theme));
    let children = &records[0].as_ref().expect("WPG drawing").children;
    assert_eq!(children.len(), 2);

    let translucent_text_slot = children[0].shape.as_ref().expect("tx1 shape");
    assert_eq!(
        translucent_text_slot.shape.fill,
        Some(Color::new(0, 0, 0)),
        "tx1 must resolve through the document theme's dk1 slot"
    );
    approx(
        translucent_text_slot
            .shape
            .opacity
            .expect("nested alpha must reach the shape"),
        0.16,
    );

    let shaded_background_slot = children[1].shape.as_ref().expect("bg1 shape");
    assert_eq!(
        shaded_background_slot.shape.fill,
        Some(Color::new(0xBC, 0xBC, 0xBC)),
        "bg1 must resolve through lt1 before its shade transform is applied"
    );
    assert_eq!(shaded_background_slot.shape.opacity, None);
}

#[test]
fn wpg_custom_geometry_preserves_normalized_subpaths() {
    let drawing = r#"<w:drawing><wp:anchor>
<wp:positionH><wp:posOffset>0</wp:posOffset></wp:positionH>
<wp:positionV><wp:posOffset>0</wp:posOffset></wp:positionV>
<a:graphic><a:graphicData><wpg:wgp><wps:wsp>
<wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="2540000" cy="1270000"/></a:xfrm>
<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>
<a:rect l="l" t="t" r="r" b="b"/>
<a:pathLst><a:path w="200" h="100">
<a:moveTo><a:pt x="200" y="0"/></a:moveTo>
<a:lnTo><a:pt x="0" y="40"/></a:lnTo>
<a:lnTo><a:pt x="0" y="100"/></a:lnTo>
<a:lnTo><a:pt x="200" y="100"/></a:lnTo>
<a:close/>
</a:path></a:pathLst></a:custGeom>
<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
<a:ln><a:noFill/></a:ln>
</wps:spPr>
</wps:wsp></wpg:wgp></a:graphicData></a:graphic>
</wp:anchor></w:drawing>"#;

    let records = scan_wpg_drawings(&body(drawing), None);
    let shape = records[0].as_ref().expect("WPG drawing").children[0]
        .shape
        .as_ref()
        .expect("WPG shape");

    let ShapeKind::Path { subpaths } = &shape.shape.kind else {
        panic!("expected custom path, got {:?}", shape.shape.kind);
    };
    assert_eq!(subpaths.len(), 1);
    assert!(subpaths[0].closed);
    assert_eq!(
        subpaths[0].vertices,
        vec![(1.0, 0.0), (0.0, 0.4), (0.0, 1.0), (1.0, 1.0)]
    );
}

#[test]
fn wpg_gradient_fill_preserves_stops_color_transforms_and_angle() {
    let drawing = r#"<w:drawing><wp:anchor>
<wp:positionH><wp:posOffset>0</wp:posOffset></wp:positionH>
<wp:positionV><wp:posOffset>0</wp:posOffset></wp:positionV>
<a:graphic><a:graphicData><wpg:wgp><wps:wsp>
<wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm>
<a:prstGeom prst="rect"/>
<a:gradFill><a:gsLst>
<a:gs pos="0"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:gs>
<a:gs pos="100000"><a:srgbClr val="FF3366"/></a:gs>
</a:gsLst><a:lin ang="1920000" scaled="0"/></a:gradFill>
<a:ln><a:noFill/></a:ln>
</wps:spPr>
</wps:wsp></wpg:wgp></a:graphicData></a:graphic>
</wp:anchor></w:drawing>"#;
    let theme = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
</a:theme>"#;

    let records = scan_wpg_drawings(&body(drawing), Some(theme));
    let shape = records[0].as_ref().expect("WPG drawing").children[0]
        .shape
        .as_ref()
        .expect("WPG shape");
    let gradient = shape
        .shape
        .gradient_fill
        .as_ref()
        .expect("WPG gradient fill");

    assert!((gradient.angle - 32.0).abs() < 0.001);
    assert_eq!(gradient.stops.len(), 2);
    assert!((gradient.stops[0].position - 0.0).abs() < 1e-9);
    assert_eq!(gradient.stops[0].color, Color::new(0x2F, 0x52, 0x8F));
    assert!((gradient.stops[1].position - 1.0).abs() < 1e-9);
    assert_eq!(gradient.stops[1].color, Color::new(0xFF, 0x33, 0x66));
}

#[test]
fn wpg_child_rotation_and_flip_survive_nested_group_scaling() {
    let drawing = r#"<w:drawing><wp:anchor>
<wp:positionH><wp:posOffset>38100</wp:posOffset></wp:positionH>
<wp:positionV><wp:posOffset>50800</wp:posOffset></wp:positionV>
<a:graphic><a:graphicData><wpg:wgp>
<wpg:grpSpPr><a:xfrm>
<a:off x="127000" y="254000"/><a:ext cx="2540000" cy="1270000"/>
<a:chOff x="0" y="0"/><a:chExt cx="1270000" cy="635000"/>
</a:xfrm></wpg:grpSpPr>
<wpg:grpSp><wpg:grpSpPr><a:xfrm>
<a:off x="127000" y="127000"/><a:ext cx="1270000" cy="635000"/>
<a:chOff x="0" y="0"/><a:chExt cx="635000" cy="635000"/>
</a:xfrm></wpg:grpSpPr>
<wps:wsp><wps:spPr>
<a:xfrm rot="2700000" flipH="1">
<a:off x="127000" y="63500"/><a:ext cx="254000" cy="127000"/>
</a:xfrm>
<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>
<a:rect l="l" t="t" r="r" b="b"/>
<a:pathLst><a:path w="200" h="100">
<a:moveTo><a:pt x="200" y="0"/></a:moveTo>
<a:lnTo><a:pt x="0" y="40"/></a:lnTo>
<a:lnTo><a:pt x="0" y="100"/></a:lnTo>
<a:close/>
</a:path></a:pathLst></a:custGeom>
<a:gradFill><a:gsLst>
<a:gs pos="0"><a:srgbClr val="4472C4"/></a:gs>
<a:gs pos="100000"><a:srgbClr val="FF3366"/></a:gs>
</a:gsLst><a:lin ang="1920000" scaled="0"/></a:gradFill>
<a:ln><a:noFill/></a:ln>
</wps:spPr></wps:wsp>
</wpg:grpSp></wpg:wgp></a:graphicData></a:graphic>
</wp:anchor></w:drawing>"#;

    let records = scan_wpg_drawings(&body(drawing), None);
    let child = &records[0].as_ref().expect("WPG drawing").children[0];
    approx(
        child
            .rotation_deg
            .expect("the WPG text overlay must share the shape rotation"),
        45.0,
    );
    let shape = child.shape.as_ref().expect("WPG shape");

    approx(shape.offset_x, 73.0);
    approx(shape.offset_y, 54.0);
    approx(shape.width, 80.0);
    approx(shape.height, 20.0);
    approx(
        shape
            .shape
            .rotation_deg
            .expect("the child rotation must survive"),
        45.0,
    );
    approx(
        shape
            .shape
            .gradient_fill
            .as_ref()
            .expect("the flipped child must retain its gradient")
            .angle,
        148.0,
    );
    let ShapeKind::Path { subpaths } = &shape.shape.kind else {
        panic!("expected custom path, got {:?}", shape.shape.kind);
    };
    assert_eq!(
        subpaths[0].vertices,
        vec![(0.0, 0.0), (1.0, 0.4), (1.0, 1.0)],
        "flipH must mirror every normalized custom-geometry point"
    );
}
