//! A header or footer picture anchored against the page or its margins is
//! drawn at its `wp:anchor` offsets, outside the story's flow, and a footer's
//! lines make room for one that spans the text column (issue #2327).
use super::*;
use std::io::Cursor;

const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

/// A4 with 72pt margins, a 35.4pt footer distance and a 36pt header distance.
const PAGE_WIDTH_PT: f64 = 595.3;
const PAGE_HEIGHT_PT: f64 = 841.9;
const MARGIN_PT: f64 = 72.0;
const FOOTER_DISTANCE_PT: f64 = 35.4;

const EMU_PER_PT: i64 = 12_700;

/// Where a test anchor sits on one axis.
#[derive(Clone, Copy)]
enum Position {
    Offset(&'static str, i64),
    Align(&'static str, &'static str),
}

impl Position {
    fn markup(self, element: &str) -> String {
        match self {
            Position::Offset(relative_from, emu) => format!(
                r#"<wp:{element} relativeFrom="{relative_from}"><wp:posOffset>{emu}</wp:posOffset></wp:{element}>"#
            ),
            Position::Align(relative_from, align) => format!(
                r#"<wp:{element} relativeFrom="{relative_from}"><wp:align>{align}</wp:align></wp:{element}>"#
            ),
        }
    }
}

struct Anchor {
    horizontal: Position,
    vertical: Position,
    /// `wp:extent`, in points.
    size_pt: (i64, i64),
    wrap: &'static str,
    behind_doc: bool,
    /// `distT`, `distB` and `distL`/`distR`, in points.
    distance_pt: (i64, i64, i64),
    /// A `wp:extent` height in EMU that is not a whole point.
    size_emu_height: Option<i64>,
}

impl Anchor {
    fn new(horizontal: Position, vertical: Position, size_pt: (i64, i64)) -> Self {
        Self {
            horizontal,
            vertical,
            size_pt,
            wrap: r#"<wp:wrapSquare wrapText="bothSides"/>"#,
            behind_doc: false,
            distance_pt: (0, 0, 9),
            size_emu_height: None,
        }
    }

    fn wrap(mut self, wrap: &'static str) -> Self {
        self.wrap = wrap;
        self
    }

    fn behind(mut self) -> Self {
        self.behind_doc = true;
        self
    }

    fn distance(mut self, top: i64, bottom: i64, sides: i64) -> Self {
        self.distance_pt = (top, bottom, sides);
        self
    }

    /// One run holding the anchored picture of the package's one image, in
    /// Arial of `half_points`.
    fn run(&self, id: u32, half_points: u32) -> String {
        let cx: i64 = self.size_pt.0 * EMU_PER_PT;
        let cy: i64 = self.size_emu_height.unwrap_or(self.size_pt.1 * EMU_PER_PT);
        let (top, bottom, sides) = self.distance_pt;
        format!(
            r#"<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="{half_points}"/></w:rPr><w:drawing><wp:anchor distT="{dist_t}" distB="{dist_b}" distL="{dist_s}" distR="{dist_s}"
                simplePos="0" relativeHeight="{id}" behindDoc="{behind}" locked="0" layoutInCell="1" allowOverlap="1">
                <wp:simplePos x="0" y="0"/>{horizontal}{vertical}
                <wp:extent cx="{cx}" cy="{cy}"/>{wrap}
                <wp:docPr id="{id}" name="Picture {id}"/>
                <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:pic>
                        <pic:nvPicPr><pic:cNvPr id="{id}" name="image1.bmp"/><pic:cNvPicPr/></pic:nvPicPr>
                        <pic:blipFill><a:blip r:embed="rIdImage1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                        <pic:spPr>
                            <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
                            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                            <a:ln><a:noFill/></a:ln>
                        </pic:spPr>
                    </pic:pic>
                </a:graphicData></a:graphic>
            </wp:anchor></w:drawing></w:r>"#,
            dist_t = top * EMU_PER_PT,
            dist_b = bottom * EMU_PER_PT,
            dist_s = sides * EMU_PER_PT,
            behind = u8::from(self.behind_doc),
            horizontal = self.horizontal.markup("positionH"),
            vertical = self.vertical.markup("positionV"),
            wrap = self.wrap,
        )
    }
}

fn text_paragraph(text: &str) -> String {
    sized_text_paragraph(text, 22)
}

/// A paragraph of one Arial run of `half_points`.
fn sized_text_paragraph(text: &str, half_points: u32) -> String {
    format!(
        r#"<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="{half_points}"/></w:rPr><w:t>{text}</w:t></w:r></w:p>"#
    )
}

/// A paragraph holding only `anchors`, with an 11pt Arial mark.
fn anchor_paragraph(anchors: &[&Anchor]) -> String {
    sized_anchor_paragraph(anchors, 22)
}

/// A paragraph holding only `anchors`, its mark and runs in Arial of
/// `half_points`.
fn sized_anchor_paragraph(anchors: &[&Anchor], half_points: u32) -> String {
    let runs: String = anchors
        .iter()
        .zip(1..)
        .map(|(anchor, id)| anchor.run(id, half_points))
        .collect();
    format!(
        r#"<w:p><w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="{half_points}"/></w:rPr></w:pPr>{runs}</w:p>"#
    )
}

/// A one-page package whose default header and footer parts are
/// `header`/`footer`, with relationship `rIdImage1` naming one BMP from each.
fn build_package(header: &str, footer: &str) -> Vec<u8> {
    let image: Vec<u8> = super::image_tests::make_test_bmp();
    let document: String = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document {NAMESPACES}><w:body>{body}<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>"#,
        body = text_paragraph("Body text"),
    );
    let image_relationship: &str = r#"<Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.bmp"/>"#;
    let story_relationships: String = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{image_relationship}</Relationships>"#
    );
    let parts: Vec<(&str, Vec<u8>)> = vec![
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bmp" ContentType="image/bmp"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>"#
                .to_vec(),
        ),
        (
            "_rels/.rels",
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#
                .to_vec(),
        ),
        (
            "word/_rels/document.xml.rels",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{image_relationship}<Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
        ("word/document.xml", document.into_bytes()),
        ("word/media/image1.bmp", image),
        (
            "word/header1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr {NAMESPACES}>{header}</w:hdr>"#
            )
            .into_bytes(),
        ),
        (
            "word/_rels/header1.xml.rels",
            story_relationships.clone().into_bytes(),
        ),
        (
            "word/footer1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr {NAMESPACES}>{footer}</w:ftr>"#
            )
            .into_bytes(),
        ),
        (
            "word/_rels/footer1.xml.rels",
            story_relationships.into_bytes(),
        ),
    ];
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, content) in parts {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        std::io::Write::write_all(&mut zip, &content).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn parse(header: &str, footer: &str) -> Document {
    let data: Vec<u8> = build_package(header, footer);
    DocxParser
        .parse(&data, &ConvertOptions::default())
        .unwrap()
        .0
}

fn flow_page(document: &Document) -> &FlowPage {
    match &document.pages[0] {
        Page::Flow(page) => page,
        _ => panic!("Expected FlowPage"),
    }
}

/// A story picture overlay's page-space top-left corner, its size, its wrap
/// and whether it sits behind the text.
type StoryPicture = ((f64, f64), (f64, f64), WrapMode, bool);

/// Each picture overlay a story draws.
fn story_pictures(story: &HeaderFooter) -> Vec<StoryPicture> {
    story
        .shapes
        .iter()
        .filter_map(|shape| {
            let HeaderFooterShapeContent::Picture(picture) = &shape.content else {
                return None;
            };
            assert_eq!(shape.frame.horizontal_anchor, FrameAnchor::Page);
            assert_eq!(shape.frame.vertical_anchor, FrameAnchor::Page);
            Some((
                (
                    round_to_hundredths(shape.frame.x.expect("resolved x")),
                    round_to_hundredths(shape.frame.y.expect("resolved y")),
                ),
                (shape.width, shape.height),
                picture.wrap_mode,
                shape.behind_text,
            ))
        })
        .collect()
}

fn inline_picture_count(story: &HeaderFooter) -> usize {
    story
        .paragraphs
        .iter()
        .flat_map(|paragraph| &paragraph.elements)
        .filter(|element| matches!(element, HFInline::Image(_)))
        .count()
}

fn round_to_hundredths(value: f64) -> f64 {
    (value * 100.0).round() / 100.0
}

#[test]
fn page_and_margin_anchored_footer_pictures_are_drawn_at_their_anchors() {
    let banner = Anchor::new(
        Position::Offset("margin", -25 * EMU_PER_PT),
        Position::Offset("margin", 690 * EMU_PER_PT),
        (500, 40),
    )
    .distance(1, 2, 9);
    let badge = Anchor::new(
        Position::Align("page", "center"),
        Position::Offset("page", 4 * EMU_PER_PT),
        (100, 50),
    )
    .wrap("<wp:wrapNone/>")
    .behind();
    let column_logo = Anchor::new(
        Position::Offset("column", 30 * EMU_PER_PT),
        Position::Offset("page", 790 * EMU_PER_PT),
        (60, 20),
    )
    .wrap("<wp:wrapTopAndBottom/>");
    let document = parse(
        &text_paragraph("Header"),
        &[
            text_paragraph("Page 1"),
            anchor_paragraph(&[&banner, &badge]),
            anchor_paragraph(&[&column_logo]),
        ]
        .concat(),
    );
    let footer: &HeaderFooter = flow_page(&document).footer.as_ref().expect("footer");

    assert_eq!(
        story_pictures(footer),
        vec![
            (
                (MARGIN_PT - 25.0, MARGIN_PT + 690.0),
                (500.0, 40.0),
                WrapMode::Square,
                false
            ),
            (
                (round_to_hundredths((PAGE_WIDTH_PT - 100.0) / 2.0), 4.0),
                (100.0, 50.0),
                WrapMode::Behind,
                true
            ),
            (
                (MARGIN_PT + 30.0, 790.0),
                (60.0, 20.0),
                WrapMode::TopAndBottom,
                false
            ),
        ],
        "each picture sits at its own anchor, against the page or the margins"
    );
    let HeaderFooterShapeContent::Picture(banner_picture) = &footer
        .shapes
        .iter()
        .find(|shape| matches!(shape.content, HeaderFooterShapeContent::Picture(_)))
        .unwrap()
        .content
    else {
        unreachable!()
    };
    assert_eq!(
        banner_picture.wrap_distance,
        Insets {
            top: 1.0,
            right: 9.0,
            bottom: 2.0,
            left: 9.0
        }
    );
    assert_eq!(
        inline_picture_count(footer),
        0,
        "no anchored picture is laid out in the story's lines"
    );
    assert_eq!(
        footer.paragraphs.len(),
        3,
        "the anchor paragraphs stay in the story as empty lines"
    );
}

#[test]
fn a_paragraph_relative_story_picture_stays_in_the_story() {
    // TODO(#2414): a picture positioned against its paragraph is still laid
    // out in the story's lines; only page- and margin-relative ones move.
    let logo = Anchor::new(
        Position::Offset("column", 0),
        Position::Offset("paragraph", 12 * EMU_PER_PT),
        (60, 20),
    );
    let document = parse(&anchor_paragraph(&[&logo]), &text_paragraph("Footer"));
    let header: &HeaderFooter = flow_page(&document).header.as_ref().expect("header");

    assert!(story_pictures(header).is_empty());
    assert_eq!(inline_picture_count(header), 1);
}

#[test]
fn a_page_anchored_header_picture_leaves_the_header_flow() {
    let logo = Anchor::new(
        Position::Offset("page", 40 * EMU_PER_PT),
        Position::Offset("margin", -50 * EMU_PER_PT),
        (80, 30),
    )
    .wrap("<wp:wrapNone/>");
    let document = parse(
        &[anchor_paragraph(&[&logo]), text_paragraph("Header")].concat(),
        &text_paragraph("Footer"),
    );
    let header: &HeaderFooter = flow_page(&document).header.as_ref().expect("header");

    assert_eq!(
        story_pictures(header),
        vec![(
            (40.0, MARGIN_PT - 50.0),
            (80.0, 30.0),
            WrapMode::None,
            false
        )]
    );
    assert_eq!(inline_picture_count(header), 0);
}

#[cfg(not(target_arch = "wasm32"))]
mod rendered {
    use super::*;

    /// The compiled first page's picture top-left corners and the baseline of
    /// the run reading `Page 1`.
    fn compile(document: &Document) -> (Vec<(f64, f64)>, f64) {
        let output = crate::render::typst_gen::generate_typst(document).unwrap();
        let pictures: Vec<(f64, f64)> =
            crate::render::pdf::compiled_image_boxes(&output.source, &output.images, 0)
                .unwrap_or_else(|error| panic!("compile failed: {error}\n{}", output.source))
                .iter()
                .map(|picture| {
                    (
                        round_to_hundredths(picture.corners[0].0),
                        round_to_hundredths(picture.corners[0].1),
                    )
                })
                .collect();
        let baseline: f64 =
            crate::render::pdf::compiled_text_runs_with_images(&output.source, &output.images, 0)
                .unwrap()
                .iter()
                .find(|run| run.text.contains("Page 1"))
                .expect("the footer line is placed")
                .baseline_pt;
        (pictures, baseline)
    }

    /// The footer `Page 1` line, then a paragraph holding only `anchor`.
    fn footer_with(anchor: &Anchor) -> Document {
        parse(
            &text_paragraph("Header"),
            &[text_paragraph("Page 1"), anchor_paragraph(&[anchor])].concat(),
        )
    }

    /// How far the footer's `Page 1` line sits above where it does beside a
    /// picture that wraps nothing.
    fn lift_pt(anchor: Anchor) -> f64 {
        let unwrapped = Anchor {
            wrap: "<wp:wrapNone/>",
            ..anchor
        };
        let (_, resting) = compile(&footer_with(&unwrapped));
        let (_, lifted) = compile(&footer_with(&anchor));
        round_to_hundredths(resting - lifted)
    }

    /// Where the footer story's bottom edge rests: `w:footer` above the page
    /// bottom.
    const STORY_BOTTOM_PT: f64 = PAGE_HEIGHT_PT - FOOTER_DISTANCE_PT;

    #[test]
    fn a_margin_anchored_footer_picture_is_painted_at_its_anchor() {
        let banner = Anchor::new(
            Position::Offset("margin", -25 * EMU_PER_PT),
            Position::Offset("margin", 690 * EMU_PER_PT),
            (500, 40),
        );
        let (pictures, _) = compile(&footer_with(&banner));
        assert_eq!(pictures, vec![(MARGIN_PT - 25.0, MARGIN_PT + 690.0)]);
    }

    /// Word sets a footer's lines above a picture that leaves no room beside
    /// them: the story's bottom edge moves up to the picture's top, less
    /// `distT`. The docxide-pdf `cyprus_ucits_marketing_registry` page number
    /// sits two lines above its full-width footer picture (issue #2327).
    #[test]
    fn a_footer_rises_above_a_picture_that_spans_its_lines() {
        let spanning = |top_pt: i64, dist_t_pt: i64| {
            Anchor::new(
                Position::Offset("margin", -25 * EMU_PER_PT),
                Position::Offset("margin", (top_pt - 72) * EMU_PER_PT),
                (500, 40),
            )
            .distance(dist_t_pt, 0, 9)
        };
        for (top_pt, dist_t_pt) in [(762, 0), (772, 3), (790, 1)] {
            let expected: f64 = round_to_hundredths(STORY_BOTTOM_PT - (top_pt - dist_t_pt) as f64);
            assert_eq!(
                lift_pt(spanning(top_pt, dist_t_pt)),
                expected,
                "picture top {top_pt}pt, distT {dist_t_pt}pt"
            );
        }

        let narrow_top_and_bottom = Anchor::new(
            Position::Offset("margin", 100 * EMU_PER_PT),
            Position::Offset("page", 780 * EMU_PER_PT),
            (60, 20),
        )
        .wrap("<wp:wrapTopAndBottom/>");
        assert_eq!(
            lift_pt(narrow_top_and_bottom),
            round_to_hundredths(STORY_BOTTOM_PT - 780.0),
            "top-and-bottom wrap keeps text off both sides however narrow the picture"
        );
    }

    fn page_1_baseline(footer: &str) -> f64 {
        compile(&parse(&text_paragraph("Header"), footer)).1
    }

    /// A paragraph whose only content is an anchored picture is still a line
    /// of the story, as an empty one is in Word: the docxide-pdf
    /// `cyprus_ucits_marketing_registry` page number sits two lines above its
    /// footer picture, not one (issue #2327).
    #[test]
    fn an_anchor_only_footer_paragraph_keeps_its_line() {
        let far_away = Anchor::new(
            Position::Offset("page", 0),
            Position::Offset("page", 0),
            (20, 10),
        )
        .wrap("<wp:wrapNone/>");
        for half_points in [22, 40] {
            let anchored: f64 = page_1_baseline(
                &[
                    text_paragraph("Page 1"),
                    sized_anchor_paragraph(&[&far_away], half_points),
                ]
                .concat(),
            );
            let lettered: f64 = page_1_baseline(
                &[
                    text_paragraph("Page 1"),
                    sized_text_paragraph("Lower", half_points),
                ]
                .concat(),
            );
            assert_eq!(
                round_to_hundredths(anchored),
                round_to_hundredths(lettered),
                "a {half_points}-half-point anchor line under the page number"
            );
        }
    }

    /// The story's lines cross a picture by Word's line boxes, which reach a
    /// face's ascent above the first baseline, not only its cap height.
    #[test]
    fn a_footer_rises_above_a_picture_its_word_line_box_reaches() {
        let Some((_, _, pitch_em)) = crate::render::pdf::font_line_metrics_em("Arial") else {
            return; // the face is unavailable on this runner
        };
        // `Page 1` and the anchor paragraph are two 11pt Arial lines. A
        // package without a styles part takes Word's built-in 1.158-line
        // multiple (issue #2215), which Word's footer boxes stack at
        // (issue #2062).
        let probe: Document = footer_with(&Anchor::new(
            Position::Offset("page", 0),
            Position::Offset("page", 0),
            (20, 10),
        ));
        let multiple: f64 = match flow_page(&probe).footer.as_ref().unwrap().paragraphs[0]
            .style
            .line_spacing
        {
            Some(crate::ir::LineSpacing::Proportional(factor)) => factor,
            _ => 1.0,
        };
        let word_top: f64 = STORY_BOTTOM_PT - 2.0 * multiple * pitch_em * 11.0;
        let ending_at = |bottom_pt: f64| {
            let height_emu: i64 = (bottom_pt * EMU_PER_PT as f64).round() as i64 - 740 * EMU_PER_PT;
            let mut anchor = Anchor::new(
                Position::Offset("margin", -25 * EMU_PER_PT),
                Position::Offset("page", 740 * EMU_PER_PT),
                (500, 0),
            );
            anchor.size_emu_height = Some(height_emu);
            anchor
        };
        assert_eq!(
            lift_pt(ending_at(word_top + 0.5)),
            round_to_hundredths(STORY_BOTTOM_PT - 740.0),
            "a picture ending inside the first line's ascent"
        );
        assert_eq!(
            lift_pt(ending_at(word_top - 0.5)),
            0.0,
            "a picture ending above the first line"
        );
    }

    #[test]
    fn a_footer_stays_put_beside_a_picture_clear_of_its_lines() {
        let at_top = |top_pt: i64, wrap: &'static str| {
            Anchor::new(
                Position::Offset("margin", -25 * EMU_PER_PT),
                Position::Offset("page", top_pt * EMU_PER_PT),
                (500, 40),
            )
            .wrap(wrap)
        };
        let square: &'static str = r#"<wp:wrapSquare wrapText="bothSides"/>"#;
        assert_eq!(lift_pt(at_top(600, square)), 0.0, "above the story");
        assert_eq!(lift_pt(at_top(810, square)), 0.0, "below the story");
        assert_eq!(
            lift_pt(at_top(770, square).behind()),
            0.0,
            "behind the text"
        );
        // Word sets the story's lines beside a picture that leaves room for
        // them, so they keep their height.
        let narrow = Anchor::new(
            Position::Offset("margin", 0),
            Position::Offset("page", 780 * EMU_PER_PT),
            (60, 20),
        );
        assert_eq!(lift_pt(narrow), 0.0, "room beside the picture");
    }
}
