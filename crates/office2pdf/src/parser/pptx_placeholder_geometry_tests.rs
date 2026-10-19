use super::*;
use std::io::Write;
use zip::write::SimpleFileOptions as FileOptions;

// ── Helpers ──────────────────────────────────────────────────────────

const LAYOUT_HEADER: &str = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#;
const LAYOUT_FOOTER: &str = "</p:spTree></p:cSld></p:sldLayout>";

const MASTER_HEADER: &str = r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#;
const MASTER_FOOTER: &str = "</p:spTree></p:cSld></p:sldMaster>";

/// A placeholder `<p:sp>` for a slide, layout, or master.
/// `ph_attrs` is the raw attribute string of `<p:ph>` (e.g. `type="title"` or `idx="1"`).
/// `xfrm_emu` is `Some((x, y, cx, cy))` for an explicit `<a:xfrm>`, or `None` to inherit.
fn make_placeholder_sp(
    ph_attrs: &str,
    xfrm_emu: Option<(i64, i64, i64, i64)>,
    text: &str,
) -> String {
    let sp_pr: String = match xfrm_emu {
        Some((x, y, cx, cy)) => format!(
            r#"<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr>"#
        ),
        None => "<p:spPr/>".to_string(),
    };
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph {ph_attrs}/></p:nvPr></p:nvSpPr>{sp_pr}<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

fn make_slide_with_shapes(shapes: &[String]) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
    );
    for shape in shapes {
        xml.push_str(shape);
    }
    xml.push_str("</p:spTree></p:cSld></p:sld>");
    xml
}

fn make_layout_with_shapes(shapes: &[String]) -> String {
    let mut xml = String::from(LAYOUT_HEADER);
    for shape in shapes {
        xml.push_str(shape);
    }
    xml.push_str(LAYOUT_FOOTER);
    xml
}

fn make_master_with_shapes(shapes: &[String]) -> String {
    let mut xml = String::from(MASTER_HEADER);
    for shape in shapes {
        xml.push_str(shape);
    }
    xml.push_str(MASTER_FOOTER);
    xml
}

fn parse_document(data: &[u8]) -> Document {
    let parser = PptxParser;
    let (doc, _warnings) = parser.parse(data, &ConvertOptions::default()).unwrap();
    doc
}

/// Find the text box element containing the given text.
fn find_text_box_with_text<'a>(page: &'a FixedPage, needle: &str) -> &'a FixedElement {
    page.elements
        .iter()
        .find(|element| {
            if let FixedElementKind::TextBox(text_box) = &element.kind {
                text_box.content.iter().any(|block| match block {
                    Block::Paragraph(paragraph) => {
                        paragraph.runs.iter().any(|run| run.text.contains(needle))
                    }
                    _ => false,
                })
            } else {
                false
            }
        })
        .unwrap_or_else(|| panic!("no text box containing {needle:?}"))
}

fn assert_geometry(element: &FixedElement, x_emu: i64, y_emu: i64, cx_emu: i64, cy_emu: i64) {
    let expected: [f64; 4] = [
        emu_to_pt(x_emu),
        emu_to_pt(y_emu),
        emu_to_pt(cx_emu),
        emu_to_pt(cy_emu),
    ];
    let actual: [f64; 4] = [element.x, element.y, element.width, element.height];
    for (index, (value, want)) in actual.iter().zip(expected.iter()).enumerate() {
        assert!(
            (value - want).abs() < 0.01,
            "geometry component {index}: got {value}, want {want} (element: x={} y={} w={} h={})",
            element.x,
            element.y,
            element.width,
            element.height
        );
    }
}

// ── Slide → layout inheritance ───────────────────────────────────────

#[test]
fn test_title_placeholder_inherits_layout_geometry() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(r#"type="title""#, None, "Hello")]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp(
        r#"type="title""#,
        Some((457_200, 274_638, 8_229_600, 1_143_000)),
        "Layout title",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let element = find_text_box_with_text(page, "Hello");
    assert_geometry(element, 457_200, 274_638, 8_229_600, 1_143_000);
}

#[test]
fn test_ctr_title_placeholder_matches_layout_title_family() {
    // A slide `title` placeholder must match a layout `ctrTitle` placeholder
    // (and vice versa): both belong to the title family.
    let slide = make_slide_with_shapes(&[make_placeholder_sp(r#"type="title""#, None, "Hello")]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp(
        r#"type="ctrTitle""#,
        Some((685_800, 2_130_425, 7_772_400, 1_470_025)),
        "Centered layout title",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let element = find_text_box_with_text(page, "Hello");
    assert_geometry(element, 685_800, 2_130_425, 7_772_400, 1_470_025);
}

#[test]
fn test_body_placeholders_match_layout_by_idx() {
    let slide = make_slide_with_shapes(&[
        make_placeholder_sp(r#"type="body" idx="1""#, None, "Left"),
        make_placeholder_sp(r#"type="body" idx="2""#, None, "Right"),
    ]);
    let layout = make_layout_with_shapes(&[
        make_placeholder_sp(
            r#"type="body" idx="1""#,
            Some((457_200, 1_600_200, 4_038_600, 4_525_963)),
            "Layout left",
        ),
        make_placeholder_sp(
            r#"type="body" idx="2""#,
            Some((4_648_200, 1_600_200, 4_038_600, 4_525_963)),
            "Layout right",
        ),
    ]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let left = find_text_box_with_text(page, "Left");
    assert_geometry(left, 457_200, 1_600_200, 4_038_600, 4_525_963);
    let right = find_text_box_with_text(page, "Right");
    assert_geometry(right, 4_648_200, 1_600_200, 4_038_600, 4_525_963);
}

// ── Layout → master fallback ─────────────────────────────────────────

#[test]
fn test_layout_placeholder_without_geometry_falls_back_to_master() {
    // The layout declares the placeholder but omits <a:xfrm>; geometry must
    // come from the master's matching placeholder.
    let slide =
        make_slide_with_shapes(&[make_placeholder_sp(r#"type="body" idx="1""#, None, "Hello")]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp(
        r#"type="body" idx="1""#,
        None,
        "Layout body",
    )]);
    let master = make_master_with_shapes(&[make_placeholder_sp(
        r#"type="body" idx="1""#,
        Some((457_200, 1_600_200, 8_229_600, 4_525_963)),
        "Master body",
    )]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let element = find_text_box_with_text(page, "Hello");
    assert_geometry(element, 457_200, 1_600_200, 8_229_600, 4_525_963);
}

#[test]
fn test_subtitle_placeholder_falls_back_to_master_body() {
    // `subTitle` has no direct master counterpart; it must normalize to the
    // master `body` placeholder when the layout provides no geometry.
    let slide = make_slide_with_shapes(&[make_placeholder_sp(
        r#"type="subTitle" idx="1""#,
        None,
        "Hello",
    )]);
    let layout = make_layout_with_shapes(&[]);
    let master = make_master_with_shapes(&[make_placeholder_sp(
        r#"type="body" idx="1""#,
        Some((1_371_600, 3_886_200, 6_400_800, 1_752_600)),
        "Master body",
    )]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let element = find_text_box_with_text(page, "Hello");
    assert_geometry(element, 1_371_600, 3_886_200, 6_400_800, 1_752_600);
}

#[test]
fn test_footer_placeholder_matches_master_by_type_despite_idx_mismatch() {
    // Real decks give the footer different idx values on each level
    // (layout idx="11", master idx="4"); the footer family matches by type.
    let slide = make_slide_with_shapes(&[make_placeholder_sp(
        r#"type="ftr" sz="quarter" idx="11""#,
        None,
        "Prislista",
    )]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp(
        r#"type="ftr" sz="quarter" idx="11""#,
        None,
        "Layout footer",
    )]);
    let master = make_master_with_shapes(&[make_placeholder_sp(
        r#"type="ftr" sz="quarter" idx="4""#,
        Some((3_124_200, 6_356_350, 2_895_600, 365_125)),
        "Master footer",
    )]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let element = find_text_box_with_text(page, "Prislista");
    assert_geometry(element, 3_124_200, 6_356_350, 2_895_600, 365_125);
}

// ── Explicit slide geometry wins ─────────────────────────────────────

#[test]
fn test_placeholder_with_explicit_xfrm_keeps_own_geometry() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(
        r#"type="title""#,
        Some((914_400, 914_400, 3_657_600, 914_400)),
        "Hello",
    )]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp(
        r#"type="title""#,
        Some((457_200, 274_638, 8_229_600, 1_143_000)),
        "Layout title",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let element = find_text_box_with_text(page, "Hello");
    assert_geometry(element, 914_400, 914_400, 3_657_600, 914_400);
}

// ── Picture placeholder ──────────────────────────────────────────────

/// Build a PPTX with one slide (with an image), one layout, and one master.
fn build_test_pptx_with_layout_master_and_image(
    slide_xml: &str,
    layout_xml: &str,
    master_xml: &str,
    image_extension: &str,
    image_bytes: &[u8],
) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let opts = FileOptions::default();

    let ct = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="{image_extension}" ContentType="image/{image_extension}"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>"#
    );
    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(ct.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
    )
    .unwrap();

    let pres = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldSz cx="{SLIDE_CX}" cy="{SLIDE_CY}"/><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>"#,
    );
    zip.start_file("ppt/presentation.xml", opts).unwrap();
    zip.write_all(pres.as_bytes()).unwrap();

    zip.start_file("ppt/_rels/presentation.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/></Relationships>"#,
    )
    .unwrap();

    zip.start_file("ppt/slides/slide1.xml", opts).unwrap();
    zip.write_all(slide_xml.as_bytes()).unwrap();

    zip.start_file("ppt/slides/_rels/slide1.xml.rels", opts)
        .unwrap();
    zip.write_all(
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.{image_extension}"/></Relationships>"#
        )
        .as_bytes(),
    )
    .unwrap();

    zip.start_file("ppt/slideLayouts/slideLayout1.xml", opts)
        .unwrap();
    zip.write_all(layout_xml.as_bytes()).unwrap();

    zip.start_file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", opts)
        .unwrap();
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>"#,
    )
    .unwrap();

    zip.start_file("ppt/slideMasters/slideMaster1.xml", opts)
        .unwrap();
    zip.write_all(master_xml.as_bytes()).unwrap();

    zip.start_file(format!("ppt/media/image1.{image_extension}"), opts)
        .unwrap();
    zip.write_all(image_bytes).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn test_picture_placeholder_inherits_layout_geometry() {
    let pic = r#"<p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture"/><p:cNvPicPr/><p:nvPr><p:ph type="pic" idx="1"/></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic>"#;
    let slide = make_slide_with_shapes(&[pic.to_string()]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp(
        r#"type="pic" idx="1""#,
        Some((2_286_000, 1_143_000, 4_572_000, 3_429_000)),
        "Layout picture caption",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master_and_image(
        &slide,
        &layout,
        &master,
        "bmp",
        &image_tests::make_test_bmp(),
    );

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let element = page
        .elements
        .iter()
        .find(|element| matches!(element.kind, FixedElementKind::Image(_)))
        .expect("no image element on page");
    assert_geometry(element, 2_286_000, 1_143_000, 4_572_000, 3_429_000);
}

// ── Picture placeholder shape geometry (issue #1373) ─────────────────

/// A layout picture placeholder that states its own outline: the "crop to
/// shape" a template applies to every photo dropped into it. Mirrors a real
/// template's `<a:noFill/>` and invisible outline.
fn make_shaped_picture_placeholder_sp(
    ph_attrs: &str,
    (x, y, cx, cy): (i64, i64, i64, i64),
    geometry_xml: &str,
) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Picture Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph {ph_attrs}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>{geometry_xml}<a:noFill/><a:ln w="387350"><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Click icon to add picture</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

/// A slide picture bound to a placeholder, with the given `<p:spPr>` body.
fn make_placeholder_pic(ph_attrs: &str, sp_pr_body: &str) -> String {
    format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture"/><p:cNvPicPr/><p:nvPr><p:ph {ph_attrs}/></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rId10"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>{sp_pr_body}</p:spPr></p:pic>"#
    )
}

/// An opaque 16x16 PNG: large enough that its corners lie outside any
/// inscribed ellipse or centred triangle while its centre lies inside.
fn make_opaque_png() -> Vec<u8> {
    let mut png = std::io::Cursor::new(Vec::new());
    image::DynamicImage::ImageRgba8(image::RgbaImage::from_pixel(
        16,
        16,
        image::Rgba([200, 120, 40, 255]),
    ))
    .write_to(&mut png, image::ImageFormat::Png)
    .unwrap();
    png.into_inner()
}

fn first_image(page: &FixedPage) -> (&FixedElement, &ImageData) {
    let element = page
        .elements
        .iter()
        .find(|element| matches!(element.kind, FixedElementKind::Image(_)))
        .expect("no image element on page");
    (element, image_tests::get_image(element))
}

/// Alpha at the top-left corner and at the centre of a decoded image.
fn corner_and_centre_alpha(image: &ImageData) -> (u8, u8) {
    let decoded = image::load_from_memory(&image.data)
        .expect("image data decodes")
        .into_rgba8();
    let (width, height) = decoded.dimensions();
    (
        decoded.get_pixel(0, 0)[3],
        decoded.get_pixel(width / 2, height / 2)[3],
    )
}

const ELLIPSE_PRSTGEOM: &str = r#"<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>"#;

/// Issue #1373: a photo dropped into a template's circular picture placeholder
/// carries `<p:spPr/>` and inherits the layout's `ellipse` exactly as it
/// inherits the layout's position. Painting the whole raster leaves the
/// photo's rectangular canvas visible around the circle.
#[test]
fn test_picture_placeholder_inherits_the_layout_ellipse_clip() {
    let slide = make_slide_with_shapes(&[make_placeholder_pic(r#"type="pic" idx="14""#, "")]);
    let layout = make_layout_with_shapes(&[make_shaped_picture_placeholder_sp(
        r#"type="pic" idx="14""#,
        (1_772_412, 2_219_248, 2_414_016, 2_414_016),
        ELLIPSE_PRSTGEOM,
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master_and_image(
        &slide,
        &layout,
        &master,
        "png",
        &make_opaque_png(),
    );

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let (element, image) = first_image(page);
    assert_geometry(element, 1_772_412, 2_219_248, 2_414_016, 2_414_016);
    // An ellipse is baked into the alpha channel, never a Typst radius.
    assert_eq!(image.clip_shape, None);
    assert_eq!(image.format, ImageFormat::Png);
    let (corner, centre) = corner_and_centre_alpha(image);
    assert_eq!(
        corner, 0,
        "corner outside the inherited ellipse must be transparent"
    );
    assert_eq!(
        centre, 255,
        "centre inside the inherited ellipse must stay opaque"
    );
}

/// DrawingML inherits each shape property on its own: a picture that restates
/// its `<a:xfrm>` but stays silent on geometry still takes the layout's clip.
#[test]
fn test_picture_with_its_own_xfrm_still_inherits_the_layout_clip() {
    let slide = make_slide_with_shapes(&[make_placeholder_pic(
        r#"type="pic" idx="17""#,
        r#"<a:xfrm><a:off x="8005572" y="2196083"/><a:ext cx="2414016" cy="2414016"/></a:xfrm>"#,
    )]);
    let layout = make_layout_with_shapes(&[make_shaped_picture_placeholder_sp(
        r#"type="pic" idx="17""#,
        (100, 200, 1_000_000, 1_000_000),
        ELLIPSE_PRSTGEOM,
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master_and_image(
        &slide,
        &layout,
        &master,
        "png",
        &make_opaque_png(),
    );

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let (element, image) = first_image(page);
    assert_geometry(element, 8_005_572, 2_196_083, 2_414_016, 2_414_016);
    let (corner, centre) = corner_and_centre_alpha(image);
    assert_eq!(
        corner, 0,
        "the slide's own xfrm does not cancel the inherited clip"
    );
    assert_eq!(centre, 255);
}

/// A picture that states its own geometry has answered; the layout's ellipse
/// must not override a slide-side rectangle.
#[test]
fn test_a_picture_declared_geometry_overrides_the_layout_clip() {
    let png = make_opaque_png();
    let slide = make_slide_with_shapes(&[make_placeholder_pic(
        r#"type="pic" idx="14""#,
        r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#,
    )]);
    let layout = make_layout_with_shapes(&[make_shaped_picture_placeholder_sp(
        r#"type="pic" idx="14""#,
        (1_772_412, 2_219_248, 2_414_016, 2_414_016),
        ELLIPSE_PRSTGEOM,
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master_and_image(&slide, &layout, &master, "png", &png);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let (_, image) = first_image(page);
    assert_eq!(image.clip_shape, None);
    assert_eq!(
        image.data, png,
        "an unclipped rectangle keeps the original bytes"
    );
}

/// A layout picture placeholder can also be an `<a:custGeom>`; that path
/// inherits along the same chain and is baked into the alpha mask.
#[test]
fn test_picture_placeholder_inherits_the_layout_custom_geometry_mask() {
    let slide = make_slide_with_shapes(&[make_placeholder_pic(r#"type="pic" idx="14""#, "")]);
    let layout = make_layout_with_shapes(&[make_shaped_picture_placeholder_sp(
        r#"type="pic" idx="14""#,
        (1_772_412, 2_219_248, 2_414_016, 2_414_016),
        TRIANGLE_CUSTGEOM,
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master_and_image(
        &slide,
        &layout,
        &master,
        "png",
        &make_opaque_png(),
    );

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let (_, image) = first_image(page);
    assert_eq!(image.format, ImageFormat::Png);
    let (corner, centre) = corner_and_centre_alpha(image);
    assert_eq!(
        corner, 0,
        "the top-left corner lies outside the inherited triangle"
    );
    assert_eq!(centre, 255, "the centre lies inside the inherited triangle");
}

// ── Layout placeholder fill ──────────────────────────────────────────

/// Like `make_placeholder_sp`, but the shape carries a solid fill — the shape
/// property a colour band behind a title lives on.
fn make_filled_placeholder_sp(
    ph_attrs: &str,
    xfrm_emu: (i64, i64, i64, i64),
    fill_hex: &str,
    text: &str,
) -> String {
    let (x, y, cx, cy) = xfrm_emu;
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph {ph_attrs}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

/// Every fill painted on the page, whatever element carries it. A filled box
/// with text renders as a text box with a background; without text it renders
/// as a shape. Both are the same ink, so the assertion is about the ink.
fn painted_fills(page: &FixedPage) -> Vec<Color> {
    page.elements
        .iter()
        .filter_map(|element| match &element.kind {
            FixedElementKind::Shape(shape) => shape.fill,
            FixedElementKind::TextBox(text_box) => text_box.fill,
            _ => None,
        })
        .collect()
}

fn element_with_fill(page: &FixedPage, fill: Color) -> &FixedElement {
    page.elements
        .iter()
        .find(|element| match &element.kind {
            FixedElementKind::Shape(shape) => shape.fill == Some(fill),
            FixedElementKind::TextBox(text_box) => text_box.fill == Some(fill),
            _ => false,
        })
        .expect("an element carries the inherited fill")
}

fn page_text(page: &FixedPage) -> String {
    fn block_text(blocks: &[Block]) -> String {
        blocks
            .iter()
            .map(|block| match block {
                Block::Paragraph(p) => p.runs.iter().map(|r| r.text.as_str()).collect::<String>(),
                _ => String::new(),
            })
            .collect()
    }
    page.elements
        .iter()
        .map(|element| match &element.kind {
            FixedElementKind::TextBox(tb) => block_text(&tb.content),
            _ => String::new(),
        })
        .collect()
}

/// A slide placeholder inherits its shape properties from the layout's
/// matching placeholder, so the colour band behind a title is declared there
/// while the slide carries only the text. The layout copy is not drawn — its
/// prompt text would come with it — so the fill has to reach the slide's own
/// shape (issue #856).
#[test]
fn test_layout_placeholder_fill_is_drawn_behind_the_slide_text() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(r#"type="title""#, None, "Hello")]);
    let layout = make_layout_with_shapes(&[make_filled_placeholder_sp(
        r#"type="title""#,
        (0, 5_367_528, 12_188_952, 1_490_472),
        "7048E8",
        "Click to edit title",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);

    assert!(
        painted_fills(page).contains(&Color::new(0x70, 0x48, 0xE8)),
        "the layout placeholder's fill must be drawn, got {:?}",
        painted_fills(page)
    );
    assert!(
        page_text(page).contains("Hello"),
        "the slide's own text must still render"
    );
    assert!(
        !page_text(page).contains("Click to edit title"),
        "the layout's prompt text must stay out of the output"
    );
}

/// The slide placeholder also inherits its geometry from the layout, so the
/// band lands on the layout's box.
#[test]
fn test_the_drawn_layout_placeholder_keeps_its_own_geometry() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(r#"type="title""#, None, "Hello")]);
    let layout = make_layout_with_shapes(&[make_filled_placeholder_sp(
        r#"type="title""#,
        (0, 5_367_528, 12_188_952, 1_490_472),
        "7048E8",
        "Click to edit title",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let band = element_with_fill(page, Color::new(0x70, 0x48, 0xE8));

    assert_geometry(band, 0, 5_367_528, 12_188_952, 1_490_472);
}

/// A layout placeholder with nothing to paint stays out of the output, so an
/// empty prompt box does not become a stray element.
#[test]
fn test_an_unfilled_layout_placeholder_still_draws_nothing() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(r#"type="title""#, None, "Hello")]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp(
        r#"type="title""#,
        Some((457_200, 274_638, 8_229_600, 1_143_000)),
        "Click to edit title",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);

    assert!(
        painted_fills(page).is_empty(),
        "nothing to paint means no fill, got {:?}",
        painted_fills(page)
    );
    assert!(!page_text(page).contains("Click to edit title"));
}

/// `<a:noFill/>` on the layout placeholder is an answer, not a gap: it must
/// end the chain rather than let the master's fill through. Without this, a
/// template's transparent subtitle boxes picked up the master's band.
#[test]
fn test_a_layout_no_fill_ends_the_inheritance_chain() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(
        r#"type="body" idx="1""#,
        None,
        "Subtitle",
    )]);
    let layout = make_layout_with_shapes(&[
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="5367528"/><a:ext cx="12188952" cy="1490472"/></a:xfrm><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Prompt</a:t></a:r></a:p></p:txBody></p:sp>"#
            .to_string(),
    ]);
    let master = make_master_with_shapes(&[make_filled_placeholder_sp(
        r#"type="body" idx="1""#,
        (0, 5_367_528, 12_188_952, 1_490_472),
        "7048E8",
        "Master prompt",
    )]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);

    assert!(
        !painted_fills(page).contains(&Color::new(0x70, 0x48, 0xE8)),
        "the layout's noFill must stop the master's band, got {:?}",
        painted_fills(page)
    );
}

/// With no `noFill` in the way, the master's fill does reach the slide — so
/// the test above is pinning the veto, not an inability to inherit at depth.
#[test]
fn test_a_master_fill_reaches_a_slide_through_a_silent_layout() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(
        r#"type="body" idx="1""#,
        None,
        "Subtitle",
    )]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp(
        r#"type="body" idx="1""#,
        Some((0, 5_367_528, 12_188_952, 1_490_472)),
        "Prompt",
    )]);
    let master = make_master_with_shapes(&[make_filled_placeholder_sp(
        r#"type="body" idx="1""#,
        (0, 5_367_528, 12_188_952, 1_490_472),
        "7048E8",
        "Master prompt",
    )]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);

    assert!(
        painted_fills(page).contains(&Color::new(0x70, 0x48, 0xE8)),
        "the master's fill must reach the slide, got {:?}",
        painted_fills(page)
    );
}

// ── Layout placeholder shape geometry ────────────────────────────────

/// Like `make_filled_placeholder_sp`, but the shape body carries the given
/// geometry XML (`<a:custGeom>…</a:custGeom>` or `<a:prstGeom …/>`) between
/// the `<a:xfrm>` and the fill — the shape of issue #1029's template panels.
fn make_shaped_filled_placeholder_sp(
    ph_attrs: &str,
    (x, y, cx, cy): (i64, i64, i64, i64),
    geometry_xml: &str,
    fill_hex: &str,
    text: &str,
) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph {ph_attrs}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>{geometry_xml}<a:solidFill><a:srgbClr val="{fill_hex}"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

/// A right-triangle `<a:custGeom>`: distinctive enough that inheriting it is
/// unmistakable from a rectangle fallback.
const TRIANGLE_CUSTGEOM: &str = r#"<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:lnTo><a:pt x="50" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>"#;

fn shape_kind_of_fill(page: &FixedPage, fill: Color) -> &ShapeKind {
    match &element_with_fill(page, fill).kind {
        FixedElementKind::Shape(shape) => &shape.kind,
        other => panic!("the fill must sit on a shape element, got {other:?}"),
    }
}

fn assert_is_the_triangle(kind: &ShapeKind) {
    let ShapeKind::Path { subpaths } = kind else {
        panic!("the inherited geometry must flatten to a path, got {kind:?}");
    };
    assert_eq!(subpaths.len(), 1, "one <a:path> yields one subpath");
    let expected: [(f64, f64); 3] = [(0.0, 1.0), (0.5, 0.0), (1.0, 1.0)];
    let vertices: &Vec<(f64, f64)> = &subpaths[0].vertices;
    for want in expected {
        assert!(
            vertices
                .iter()
                .any(|(x, y)| (x - want.0).abs() < 1e-6 && (y - want.1).abs() < 1e-6),
            "vertex {want:?} missing from {vertices:?}"
        );
    }
}

/// Issue #1029: the layout's title band is not a rectangle unless it says so.
/// The deck's blue panels are `<a:custGeom>` paths on the layout placeholder;
/// painting the inherited fill into the bounding box squared off every curved
/// corner. The path must inherit along the same chain as the fill.
#[test]
fn test_layout_placeholder_custom_geometry_shapes_the_inherited_fill() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(r#"type="title""#, None, "Hello")]);
    let layout = make_layout_with_shapes(&[make_shaped_filled_placeholder_sp(
        r#"type="title""#,
        (6_123_207, 0, 6_067_453, 6_857_999),
        TRIANGLE_CUSTGEOM,
        "32B5DF",
        "Click to edit title",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);

    assert_is_the_triangle(shape_kind_of_fill(page, Color::new(0x32, 0xB5, 0xDF)));
    assert!(page_text(page).contains("Hello"));
    assert!(!page_text(page).contains("Click to edit title"));
}

/// A slide that states its own geometry has answered; the layout's path must
/// not override it.
#[test]
fn test_a_slide_declared_geometry_overrides_the_layout_path() {
    let slide = make_slide_with_shapes(&[make_shaped_filled_placeholder_sp(
        r#"type="title""#,
        (0, 0, 2_000_000, 1_000_000),
        r#"<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>"#,
        "CD41B0",
        "Hello",
    )]);
    let layout = make_layout_with_shapes(&[make_shaped_filled_placeholder_sp(
        r#"type="title""#,
        (6_123_207, 0, 6_067_453, 6_857_999),
        TRIANGLE_CUSTGEOM,
        "32B5DF",
        "Click to edit title",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);

    assert!(
        matches!(
            shape_kind_of_fill(page, Color::new(0xCD, 0x41, 0xB0)),
            ShapeKind::Ellipse
        ),
        "the slide's own preset must win"
    );
}

/// A layout placeholder can also state a preset: a `roundRect` band must
/// inherit as a rounded rectangle, not fall back to the bounding box.
#[test]
fn test_a_layout_preset_geometry_inherits_like_a_custom_one() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(r#"type="title""#, None, "Hello")]);
    let layout = make_layout_with_shapes(&[make_shaped_filled_placeholder_sp(
        r#"type="title""#,
        (0, 0, 2_000_000, 1_000_000),
        r#"<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>"#,
        "32B5DF",
        "Click to edit title",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);

    assert!(
        matches!(
            shape_kind_of_fill(page, Color::new(0x32, 0xB5, 0xDF)),
            ShapeKind::RoundedRectangle { .. }
        ),
        "a layout roundRect must inherit as a rounded rectangle"
    );
}

/// A silent layout chains into the master for shape geometry exactly as it
/// does for fill.
#[test]
fn test_a_master_shape_geometry_reaches_a_slide_through_a_silent_layout() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(r#"type="title""#, None, "Hello")]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp(r#"type="title""#, None, "Prompt")]);
    let master = make_master_with_shapes(&[make_shaped_filled_placeholder_sp(
        r#"type="title""#,
        (6_123_207, 0, 6_067_453, 6_857_999),
        TRIANGLE_CUSTGEOM,
        "32B5DF",
        "Master prompt",
    )]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);

    assert_is_the_triangle(shape_kind_of_fill(page, Color::new(0x32, 0xB5, 0xDF)));
}

// ── Slide → layout `<a:bodyPr>` inheritance ──────────────────────────

/// A placeholder `<p:sp>` whose `<a:bodyPr>` carries the given attributes.
/// `body_pr_attrs` is the raw attribute string (e.g. `anchor="b"`).
fn make_placeholder_sp_with_body_pr(
    ph_attrs: &str,
    xfrm_emu: Option<(i64, i64, i64, i64)>,
    body_pr_attrs: &str,
    text: &str,
) -> String {
    let sp_pr: String = match xfrm_emu {
        Some((x, y, cx, cy)) => format!(
            r#"<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr>"#
        ),
        None => "<p:spPr/>".to_string(),
    };
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph {ph_attrs}/></p:nvPr></p:nvSpPr>{sp_pr}<p:txBody><a:bodyPr {body_pr_attrs}/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

fn text_box_of<'a>(page: &'a FixedPage, needle: &str) -> &'a TextBoxData {
    match &find_text_box_with_text(page, needle).kind {
        FixedElementKind::TextBox(text_box) => text_box,
        _ => unreachable!("find_text_box_with_text only matches text boxes"),
    }
}

/// The deck in issue #877 gives its footer placeholder a bare `<a:bodyPr/>`
/// on the slide while the layout placeholder declares `anchor="b"`. An
/// omitted attribute on a placeholder means "ask the layout", not "use the
/// built-in top anchor".
#[test]
fn test_placeholder_inherits_the_layout_body_pr_anchor() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(
        r#"type="ftr" sz="quarter" idx="10""#,
        None,
        "CONTOSO",
    )]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp_with_body_pr(
        r#"type="ftr" sz="quarter" idx="10""#,
        Some((795_528, 5_751_576, 4_956_048, 722_376)),
        r#"rtlCol="0" anchor="b""#,
        "Prompt",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    assert_eq!(
        text_box_of(page, "CONTOSO").vertical_align,
        TextBoxVerticalAlign::Bottom
    );
}

/// Triangulation for the test above: the same slide markup against a layout
/// that anchors to the centre must land on the centre, not on a fixed answer.
#[test]
fn test_placeholder_inherits_a_centre_anchor_too() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(r#"type="title""#, None, "Tittel")]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp_with_body_pr(
        r#"type="title""#,
        Some((7_955_280, 676_656, 3_666_744, 5_495_544)),
        r#"lIns="0" tIns="45720" rIns="0" bIns="45720" rtlCol="0" anchor="ctr""#,
        "Prompt",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    assert_eq!(
        text_box_of(page, "Tittel").vertical_align,
        TextBoxVerticalAlign::Center
    );
}

/// The slide's own `<a:bodyPr>` still wins where it states an attribute.
#[test]
fn test_a_slide_body_pr_anchor_overrides_the_layout() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp_with_body_pr(
        r#"type="title""#,
        None,
        r#"anchor="t""#,
        "Tittel",
    )]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp_with_body_pr(
        r#"type="title""#,
        Some((7_955_280, 676_656, 3_666_744, 5_495_544)),
        r#"rtlCol="0" anchor="ctr""#,
        "Prompt",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    assert_eq!(
        text_box_of(page, "Tittel").vertical_align,
        TextBoxVerticalAlign::Top
    );
}

/// Insets inherit on the same chain. The deck in issue #878 puts
/// `lIns="795528"` (62.64pt) on the layout title, where the built-in default
/// is 91440 EMU (7.2pt) — a 55.44pt difference in where the title starts.
#[test]
fn test_placeholder_inherits_the_layout_text_insets() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(r#"type="title""#, None, "Tittel")]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp_with_body_pr(
        r#"type="title""#,
        Some((0, 5_367_528, 12_188_952, 1_490_472)),
        r#"lIns="795528" tIns="338328" rtlCol="0""#,
        "Prompt",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let padding = text_box_of(page, "Tittel").padding;
    assert!(
        (padding.left - emu_to_pt(795_528)).abs() < 0.01,
        "left inset: got {}, want {}",
        padding.left,
        emu_to_pt(795_528)
    );
    assert!(
        (padding.top - emu_to_pt(338_328)).abs() < 0.01,
        "top inset: got {}, want {}",
        padding.top,
        emu_to_pt(338_328)
    );
    // `rIns`/`bIns` are absent from the layout too, so they keep the built-in
    // default rather than picking up the left/top values.
    assert!(
        (padding.right - emu_to_pt(91_440)).abs() < 0.01,
        "right inset: got {}, want the 91440 EMU default",
        padding.right
    );
}

/// A layout that states nothing passes the question on to the master.
#[test]
fn test_placeholder_body_pr_falls_back_to_the_master() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(
        r#"type="ftr" sz="quarter" idx="10""#,
        None,
        "CONTOSO",
    )]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp_with_body_pr(
        r#"type="ftr" sz="quarter" idx="10""#,
        Some((795_528, 5_751_576, 4_956_048, 722_376)),
        r#"rtlCol="0""#,
        "Prompt",
    )]);
    let master = make_master_with_shapes(&[make_placeholder_sp_with_body_pr(
        r#"type="ftr" sz="quarter" idx="3""#,
        Some((795_528, 5_751_576, 4_956_048, 722_376)),
        r#"rtlCol="0" anchor="ctr""#,
        "Master prompt",
    )]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    assert_eq!(
        text_box_of(page, "CONTOSO").vertical_align,
        TextBoxVerticalAlign::Center
    );
}

/// A plain text box is not a placeholder and inherits nothing: the deck in
/// issue #877 draws one on top of the footer placeholder, and it must keep
/// the anchor it states itself.
#[test]
fn test_a_non_placeholder_text_box_inherits_no_body_pr() {
    let text_box = r#"<p:sp><p:nvSpPr><p:cNvPr id="9" name="TextBox"/><p:cNvSpPr txBox="1"><a:spLocks/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="796918" y="5749741"/><a:ext cx="4959308" cy="721732"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Plain box</a:t></a:r></a:p></p:txBody></p:sp>"#.to_string();
    let slide = make_slide_with_shapes(&[text_box]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp_with_body_pr(
        r#"type="ftr" sz="quarter" idx="10""#,
        Some((795_528, 5_751_576, 4_956_048, 722_376)),
        r#"rtlCol="0" anchor="b""#,
        "Prompt",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    assert_eq!(
        text_box_of(page, "Plain box").vertical_align,
        TextBoxVerticalAlign::Top
    );
}

// ── Slide → layout `<a:xfrm>` rotation inheritance ───────────────────

/// A placeholder `<p:sp>` whose `<a:xfrm>` carries the given extra attributes
/// (e.g. `rot="16200000"`).
fn make_placeholder_sp_with_xfrm_attrs(
    ph_attrs: &str,
    xfrm_emu: (i64, i64, i64, i64),
    xfrm_attrs: &str,
    text: &str,
) -> String {
    let (x, y, cx, cy) = xfrm_emu;
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph {ph_attrs}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm {xfrm_attrs}><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#
    )
}

/// A slide placeholder that omits `<a:xfrm>` inherits the layout's rotation
/// along with its position and size (issue #881). The deck there rotates the
/// footer strip 270° on the layout, and the slide states no transform at all.
#[test]
fn test_placeholder_inherits_the_layout_rotation() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(
        r#"type="ftr" sz="quarter" idx="25""#,
        None,
        "CONTOSO",
    )]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp_with_xfrm_attrs(
        r#"type="ftr" sz="quarter" idx="25""#,
        (-3_049_519, 3_049_522, 6_857_999, 758_952),
        r#"rot="16200000""#,
        "Prompt",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    let rotation = text_box_of(page, "CONTOSO")
        .shape_rotation_deg
        .expect("an inherited placeholder must carry the layout's rotation");
    assert!(
        (rotation - 270.0).abs() < 0.01,
        "expected 270 degrees, got {rotation}"
    );
}

/// Triangulation: a different layout angle must come through as that angle,
/// and a layout that states none must leave the placeholder unrotated.
#[test]
fn test_placeholder_rotation_follows_the_layout_angle() {
    for (xfrm_attrs, expected) in [(r#"rot="5400000""#, Some(90.0_f64)), ("", None)] {
        let slide =
            make_slide_with_shapes(&[make_placeholder_sp(r#"type="title""#, None, "Tittel")]);
        let layout = make_layout_with_shapes(&[make_placeholder_sp_with_xfrm_attrs(
            r#"type="title""#,
            (886_968, 1_627_632, 4_416_552, 685_800),
            xfrm_attrs,
            "Prompt",
        )]);
        let master = make_master_with_shapes(&[]);
        let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

        let doc = parse_document(&data);
        let page = first_fixed_page(&doc);
        match (text_box_of(page, "Tittel").shape_rotation_deg, expected) {
            (None, None) => {}
            (Some(actual), Some(want)) => assert!(
                (actual - want).abs() < 0.01,
                "expected {want} degrees, got {actual}"
            ),
            (actual, want) => panic!("expected {want:?}, got {actual:?}"),
        }
    }
}

/// A slide placeholder that states its own `<a:xfrm>` keeps its own transform,
/// rotation included — the layout's is not merged into it.
#[test]
fn test_a_placeholder_with_its_own_xfrm_keeps_its_own_rotation() {
    let slide = make_slide_with_shapes(&[make_placeholder_sp(
        r#"type="ftr" sz="quarter" idx="25""#,
        Some((100_000, 200_000, 3_000_000, 500_000)),
        "CONTOSO",
    )]);
    let layout = make_layout_with_shapes(&[make_placeholder_sp_with_xfrm_attrs(
        r#"type="ftr" sz="quarter" idx="25""#,
        (-3_049_519, 3_049_522, 6_857_999, 758_952),
        r#"rot="16200000""#,
        "Prompt",
    )]);
    let master = make_master_with_shapes(&[]);
    let data = build_test_pptx_with_layout_master(SLIDE_CX, SLIDE_CY, &slide, &layout, &master);

    let doc = parse_document(&data);
    let page = first_fixed_page(&doc);
    assert_eq!(text_box_of(page, "CONTOSO").shape_rotation_deg, None);
}
