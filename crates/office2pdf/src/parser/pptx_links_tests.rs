use super::*;

const RELS: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout7.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide3.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/?a=1&amp;b=2" TargetMode="External"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide9.xml"/>
</Relationships>"#;

/// Slides 1, 3 and 5 print on pages 0-2; slide 9 is hidden.
fn deck() -> DeckPages {
    DeckPages {
        page_by_slide: [
            ("ppt/slides/slide1.xml", 0),
            ("ppt/slides/slide3.xml", 1),
            ("ppt/slides/slide5.xml", 2),
        ]
        .into_iter()
        .map(|(path, page)| (path.to_string(), page))
        .collect(),
        page_count: 3,
    }
}

#[test]
fn a_url_relationship_without_an_action_is_a_url() {
    let deck = deck();
    let links = PartLinks::new("ppt/slides/slide1.xml", RELS, &deck, 0);
    assert_eq!(
        links.resolve(Some("rId3"), None),
        Some(LinkTarget::Url("https://example.org/?a=1&b=2".to_string()))
    );
    // A relationship of another type is not a link, even without an action.
    assert_eq!(links.resolve(Some("rId1"), None), None);
    assert_eq!(links.resolve(Some("rId99"), None), None);
}

#[test]
fn a_slide_jump_leads_to_the_page_its_slide_prints_on() {
    let deck = deck();
    let links = PartLinks::new("ppt/slides/slide1.xml", RELS, &deck, 0);
    assert_eq!(
        links.resolve(Some("rId2"), Some("ppaction://hlinksldjump")),
        Some(LinkTarget::Page(1))
    );
    // Slide 9 prints nowhere, so the jump has no destination.
    assert_eq!(
        links.resolve(Some("rId4"), Some("ppaction://hlinksldjump")),
        None
    );
    // A jump naming a URL relationship is malformed.
    assert_eq!(
        links.resolve(Some("rId3"), Some("ppaction://hlinksldjump")),
        None
    );
}

/// A layout resolves its own relationships: `../slides/slide5.xml` from
/// `ppt/slideLayouts/` is slide 5.
#[test]
fn a_layout_resolves_its_relationships_against_its_own_folder() {
    let deck = deck();
    let layout_rels: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide5.xml"/></Relationships>"#;
    let links = PartLinks::new("ppt/slideLayouts/slideLayout7.xml", layout_rels, &deck, 0);
    assert_eq!(
        links.resolve(Some("rId5"), Some("ppaction://hlinksldjump")),
        Some(LinkTarget::Page(2))
    );
}

#[test]
fn show_jumps_move_relative_to_the_page_they_are_shown_on() {
    let deck = deck();
    let jump = |page: usize, to: &str| {
        PartLinks::new("ppt/slides/slide3.xml", RELS, &deck, page)
            .resolve(None, Some(&format!("ppaction://hlinkshowjump?jump={to}")))
    };
    assert_eq!(jump(1, "nextslide"), Some(LinkTarget::Page(2)));
    assert_eq!(jump(1, "previousslide"), Some(LinkTarget::Page(0)));
    assert_eq!(jump(1, "firstslide"), Some(LinkTarget::Page(0)));
    assert_eq!(jump(1, "lastslide"), Some(LinkTarget::Page(2)));
    // No page past either end.
    assert_eq!(jump(2, "nextslide"), None);
    assert_eq!(jump(0, "previousslide"), None);
    // Ending the show, or an unknown jump, has no PDF equivalent.
    assert_eq!(jump(1, "endshow"), None);
}

#[test]
fn actions_a_pdf_cannot_carry_resolve_to_nothing() {
    let deck = deck();
    let links = PartLinks::new("ppt/slides/slide1.xml", RELS, &deck, 0);
    for action in [
        "ppaction://macro?name=Run",
        "ppaction://program",
        "ppaction://customshow?id=0",
        "ppaction://noaction",
    ] {
        assert_eq!(links.resolve(Some("rId3"), Some(action)), None, "{action}");
    }
}

/// End to end through the PPTX parser and the renderer (issue #2473).
#[cfg(not(target_arch = "wasm32"))]
mod deck {
    use std::io::{Cursor, Write};

    use crate::config::ConvertOptions;
    use crate::ir::{
        Block, Document, FixedElementKind, FixedPage, LinkArea, LinkTarget, Page, Run,
    };
    use crate::parser::Parser as _;
    use crate::parser::pptx::PptxParser;

    const NS: &str = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main""#;
    const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const EMU_PER_PT: i64 = 12_700;

    fn rels(entries: &[(&str, &str, &str)]) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );
        for (id, kind, target) in entries {
            let mode = if *kind == "hyperlink" {
                r#" TargetMode="External""#
            } else {
                ""
            };
            xml.push_str(&format!(
                r#"<Relationship Id="{id}" Type="{REL_NS}/{kind}" Target="{target}"{mode}/>"#
            ));
        }
        xml.push_str("</Relationships>");
        xml
    }

    /// A text box at (x, y) points, w x h, with `click` inside its `p:cNvPr`
    /// and one run carrying `run_click` inside its `a:rPr`.
    fn text_box(
        id: u32,
        (x, y, w, h): (i64, i64, i64, i64),
        click: &str,
        text: &str,
        run_click: &str,
    ) -> String {
        format!(
            r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="Box {id}">{click}</p:cNvPr><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US">{run_click}</a:rPr><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"#,
            x * EMU_PER_PT,
            y * EMU_PER_PT,
            w * EMU_PER_PT,
            h * EMU_PER_PT
        )
    }

    fn slide(shapes: &str, hidden: bool) -> String {
        let show = if hidden { r#" show="0""# } else { "" };
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><p:sld {NS}{show}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree></p:cSld></p:sld>"#
        )
    }

    fn click(rid: &str, action: &str) -> String {
        let action = if action.is_empty() {
            String::new()
        } else {
            format!(r#" action="{action}""#)
        };
        format!(r#"<a:hlinkClick r:id="{rid}"{action}/>"#)
    }

    /// Four 720 x 540pt slides; slide 3 is hidden, so slides 1, 2 and 4 print
    /// on pages 0, 1 and 2. The layout carries a "last slide" button every
    /// page shows.
    fn build_deck(slides: &[(String, String)], layout_shapes: &str) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let opts = zip::write::SimpleFileOptions::default();
        let mut put = |path: &str, body: &str| {
            zip.start_file(path, opts).unwrap();
            zip.write_all(body.as_bytes()).unwrap();
        };
        let mut types = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>"#,
        );
        for index in 1..=slides.len() {
            types.push_str(&format!(r#"<Override PartName="/ppt/slides/slide{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#));
        }
        types.push_str(r#"<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/></Types>"#);
        put("[Content_Types].xml", &types);
        put(
            "_rels/.rels",
            &rels(&[("rId1", "officeDocument", "ppt/presentation.xml")]),
        );
        let ids: String = (1..=slides.len())
            .map(|index| format!(r#"<p:sldId id="{}" r:id="rId{}"/>"#, 255 + index, 1 + index))
            .collect();
        put(
            "ppt/presentation.xml",
            &format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><p:presentation {NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId100"/></p:sldMasterIdLst><p:sldSz cx="{}" cy="{}"/><p:sldIdLst>{ids}</p:sldIdLst></p:presentation>"#,
                720 * EMU_PER_PT,
                540 * EMU_PER_PT
            ),
        );
        let mut presentation_rels: Vec<(String, &str, String)> = vec![(
            "rId100".to_string(),
            "slideMaster",
            "slideMasters/slideMaster1.xml".to_string(),
        )];
        for index in 1..=slides.len() {
            presentation_rels.push((
                format!("rId{}", 1 + index),
                "slide",
                format!("slides/slide{index}.xml"),
            ));
        }
        let presentation_rels: Vec<(&str, &str, &str)> = presentation_rels
            .iter()
            .map(|(id, kind, target)| (id.as_str(), *kind, target.as_str()))
            .collect();
        put("ppt/_rels/presentation.xml.rels", &rels(&presentation_rels));
        for (index, (xml, extra_rels)) in slides.iter().enumerate() {
            let number = index + 1;
            put(&format!("ppt/slides/slide{number}.xml"), xml);
            let rels_xml = rels(&[("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")])
                .replace("</Relationships>", &format!("{extra_rels}</Relationships>"));
            put(
                &format!("ppt/slides/_rels/slide{number}.xml.rels"),
                &rels_xml,
            );
        }
        put(
            "ppt/slideLayouts/slideLayout1.xml",
            &format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><p:sldLayout {NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{layout_shapes}</p:spTree></p:cSld></p:sldLayout>"#
            ),
        );
        put(
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            &rels(&[("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")]),
        );
        put(
            "ppt/slideMasters/slideMaster1.xml",
            &format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><p:sldMaster {NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:sldMaster>"#
            ),
        );
        zip.finish().unwrap().into_inner()
    }

    fn rel(id: &str, kind: &str, target: &str) -> String {
        rels(&[(id, kind, target)])
            .trim_start_matches(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#)
            .trim_end_matches("</Relationships>")
            .to_string()
    }

    fn sample_deck() -> Document {
        let slide_one = slide(
            &[
                // A shape that jumps to slide 4, which prints on page 2.
                text_box(
                    2,
                    (72, 72, 288, 72),
                    &click("rId2", "ppaction://hlinksldjump"),
                    "Go to slide 4",
                    "",
                ),
                // A run linking out, inside a shape with no action.
                text_box(3, (72, 216, 288, 72), "", "example.org", &click("rId3", "")),
                // A jump to the hidden slide 3 has nowhere to go.
                text_box(
                    4,
                    (72, 360, 288, 72),
                    &click("rId4", "ppaction://hlinksldjump"),
                    "Hidden",
                    "",
                ),
                // A run moving to the next printed slide.
                text_box(
                    5,
                    (432, 72, 216, 72),
                    "",
                    "Next",
                    &click("", "ppaction://hlinkshowjump?jump=nextslide"),
                ),
            ]
            .concat(),
            false,
        );
        let slide_one_rels = [
            rel("rId2", "slide", "slide4.xml"),
            rel(
                "rId3",
                "hyperlink",
                "https://example.org/a?b=1&amp;c=&quot;d&quot;",
            ),
            rel("rId4", "slide", "slide3.xml"),
        ]
        .concat();
        let plain = |label: &str| slide(&text_box(2, (72, 72, 288, 72), "", label, ""), false);
        // The layout's button sits on every slide; its relationship lives in
        // the layout's own part, so it needs no slide relationship.
        let layout_button = text_box(
            9,
            (612, 468, 72, 36),
            &click("", "ppaction://hlinkshowjump?jump=lastslide"),
            "End",
            "",
        );
        build_deck(
            &[
                (slide_one, slide_one_rels),
                (plain("Slide 2"), String::new()),
                (
                    slide(&text_box(2, (72, 72, 288, 72), "", "Slide 3", ""), true),
                    String::new(),
                ),
                (plain("Slide 4"), String::new()),
            ],
            &layout_button,
        )
        .pipe(|data| {
            PptxParser
                .parse(&data, &ConvertOptions::default())
                .unwrap()
                .0
        })
    }

    trait Pipe: Sized {
        fn pipe<T>(self, f: impl FnOnce(Self) -> T) -> T {
            f(self)
        }
    }
    impl<T> Pipe for T {}

    fn fixed_pages(doc: &Document) -> Vec<&FixedPage> {
        doc.pages
            .iter()
            .map(|page| match page {
                Page::Fixed(page) => page,
                _ => panic!("a deck prints fixed pages"),
            })
            .collect()
    }

    fn runs(page: &FixedPage) -> Vec<&Run> {
        page.elements
            .iter()
            .filter_map(|element| match &element.kind {
                FixedElementKind::TextBox(text_box) => Some(&text_box.content),
                _ => None,
            })
            .flatten()
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => Some(&paragraph.runs),
                _ => None,
            })
            .flatten()
            .collect()
    }

    fn run_link<'a>(page: &'a FixedPage, text: &str) -> Option<&'a LinkTarget> {
        runs(page)
            .into_iter()
            .find(|run| run.text == text)
            .unwrap_or_else(|| panic!("no run {text:?}"))
            .href
            .as_ref()
    }

    #[test]
    fn shape_and_run_click_actions_become_page_and_url_links() {
        let doc = sample_deck();
        let pages = fixed_pages(&doc);
        assert_eq!(pages.len(), 3, "the hidden slide does not print");

        let shape_jump = LinkArea {
            x: 72.0,
            y: 72.0,
            width: 288.0,
            height: 72.0,
            target: LinkTarget::Page(2),
        };
        assert!(
            pages[0].links.contains(&shape_jump),
            "the shape jumps to slide 4's page: {:?}",
            pages[0].links
        );
        assert!(
            pages[0].links.iter().all(|area| area.y != 360.0),
            "a jump to a hidden slide has no target: {:?}",
            pages[0].links
        );
        assert_eq!(
            run_link(pages[0], "example.org"),
            Some(&LinkTarget::Url(
                "https://example.org/a?b=1&c=\"d\"".to_string()
            ))
        );
        assert_eq!(run_link(pages[0], "Next"), Some(&LinkTarget::Page(1)));
        assert_eq!(
            run_link(pages[0], "Go to slide 4"),
            None,
            "a shape link is not a run link"
        );
    }

    /// The layout's button resolves on each page it is drawn on: "last slide"
    /// is page 2 everywhere.
    #[test]
    fn a_layout_shape_links_on_every_slide_it_appears_on() {
        let doc = sample_deck();
        for (index, page) in fixed_pages(&doc).into_iter().enumerate() {
            let button = LinkArea {
                x: 612.0,
                y: 468.0,
                width: 72.0,
                height: 36.0,
                target: LinkTarget::Page(2),
            };
            assert!(
                page.links.contains(&button),
                "page {index}: {:?}",
                page.links
            );
        }
    }

    #[test]
    fn links_render_as_typst_links_with_escaped_urls() {
        let doc = sample_deck();
        let output = crate::render::typst_gen::generate_typst(&doc).unwrap();
        let source = &output.source;
        assert!(
            source.contains(r#"#link("https://example.org/a?b=1&c=\"d\"")["#),
            "the URL is escaped inside a Typst string"
        );
        assert!(
            source.contains("#place(top + left, dx: 72pt, dy: 72pt)[#link((page: 3, x: 0pt, y: 0pt))[#box(width: 288pt, height: 72pt)]]"),
            "the shape's area links to slide 4's page"
        );
        assert!(
            source.contains("#link((page: 2, x: 0pt, y: 0pt))["),
            "the next-slide run links to page 2"
        );
        // The deck compiles: no link points at a page or label it lacks.
        crate::render::pdf::compile_to_pdf(source, &output.images, None, &[], false, false)
            .expect("the deck compiles");
    }

    /// A group's own action covers everything the group draws; a table cell's
    /// run links through the slide's relationships like any other run.
    #[test]
    fn group_and_table_cell_actions_link() {
        let child = text_box(11, (100, 100, 50, 20), "", "In group", &click("rId3", ""));
        let group = format!(
            r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group">{}</p:cNvPr><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="{off}" y="{off}"/><a:ext cx="{ext}" cy="{ext}"/><a:chOff x="{off}" y="{off}"/><a:chExt cx="{ext}" cy="{ext}"/></a:xfrm></p:grpSpPr>{child}</p:grpSp>"#,
            click("rId2", "ppaction://hlinksldjump"),
            off = 100 * EMU_PER_PT,
            ext = 50 * EMU_PER_PT,
        );
        let table = format!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="20" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="{w}"/></a:tblGrid><a:tr h="{h}"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US">{}</a:rPr><a:t>cell link</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
            click("rId3", ""),
            x = 300 * EMU_PER_PT,
            y = 300 * EMU_PER_PT,
            w = 200 * EMU_PER_PT,
            h = 40 * EMU_PER_PT,
        );
        let data = build_deck(
            &[
                (
                    slide(&format!("{group}{table}"), false),
                    [
                        rel("rId2", "slide", "slide2.xml"),
                        rel("rId3", "hyperlink", "https://example.com/cell"),
                    ]
                    .concat(),
                ),
                (slide("", false), String::new()),
            ],
            "",
        );
        let doc = PptxParser
            .parse(&data, &ConvertOptions::default())
            .unwrap()
            .0;
        let page = fixed_pages(&doc)[0];
        assert!(
            page.links.contains(&LinkArea {
                x: 100.0,
                y: 100.0,
                width: 50.0,
                height: 20.0,
                target: LinkTarget::Page(1),
            }),
            "the group's action covers its child: {:?}",
            page.links
        );
        let cell_runs: Vec<&Run> = page
            .elements
            .iter()
            .filter_map(|element| match &element.kind {
                FixedElementKind::Table(table) => Some(table),
                _ => None,
            })
            .flat_map(|table| &table.rows)
            .flat_map(|row| &row.cells)
            .flat_map(|cell| &cell.content)
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => Some(&paragraph.runs),
                _ => None,
            })
            .flatten()
            .collect();
        assert_eq!(
            cell_runs.first().and_then(|run| run.href.as_ref()),
            Some(&LinkTarget::Url("https://example.com/cell".to_string()))
        );
        assert_eq!(
            run_link(page, "In group"),
            Some(&LinkTarget::Url("https://example.com/cell".to_string())),
            "a run inside a group links like any other run"
        );
    }

    /// A page link renders only into a document whose pages up to the target
    /// are all fixed, and is dropped past the end.
    #[test]
    fn a_page_link_the_document_cannot_reach_is_dropped() {
        let mut doc = sample_deck();
        let Page::Fixed(first) = &mut doc.pages[0] else {
            unreachable!()
        };
        first.links = vec![LinkArea {
            x: 0.0,
            y: 0.0,
            width: 10.0,
            height: 10.0,
            target: LinkTarget::Page(7),
        }];
        let source = crate::render::typst_gen::generate_typst(&doc)
            .unwrap()
            .source;
        assert!(!source.contains("(page: 8,"), "no link past the last page");
    }
}
