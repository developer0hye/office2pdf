use super::*;

const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const A14: &str = "http://schemas.microsoft.com/office/drawing/2010/main";

fn marker(color: &str) -> String {
    make_shape(0, 0, 914400, 914400, "rect", Some(color), None, None)
}

fn colors(content: &str) -> Vec<Color> {
    let slide = make_slide_xml(&[content.to_owned()]);
    let data = build_test_pptx(SLIDE_CX, SLIDE_CY, &[slide]);
    let (doc, warnings) = PptxParser.parse(&data, &ConvertOptions::default()).unwrap();
    assert!(warnings.is_empty(), "{warnings:?}");
    first_fixed_page(&doc)
        .elements
        .iter()
        .map(|e| get_shape(e).fill.unwrap())
        .collect()
}

#[test]
fn alternate_content_selects_only_the_first_supported_choice() {
    let xml = format!(
        r#"<compat:AlternateContent xmlns:compat="{MC}" xmlns:math="{A14}" xmlns:future="urn:future">
        <compat:Choice Requires="future">{}</compat:Choice>
        <compat:Choice Requires="math a">{}</compat:Choice>
        <compat:Choice Requires="math">{}</compat:Choice>
        <compat:Fallback>{}</compat:Fallback></compat:AlternateContent>"#,
        marker("FF0000"),
        marker("00FF00"),
        marker("0000FF"),
        marker("FFFFFF")
    );
    assert_eq!(colors(&xml), vec![Color::new(0, 255, 0)]);
}

#[test]
fn alternate_content_uses_fallback_for_unknown_or_rebound_requirements() {
    for requirement in ["future", "a14", "a missing"] {
        let xml = format!(
            r#"<mc:AlternateContent xmlns:mc="{MC}" xmlns:a14="urn:unsupported" xmlns:future="urn:future">
            <mc:Choice Requires="{requirement}">{}</mc:Choice><mc:Fallback>{}</mc:Fallback>
            </mc:AlternateContent>"#,
            marker("FF0000"),
            marker("0000FF")
        );
        assert_eq!(colors(&xml), vec![Color::new(0, 0, 255)]);
    }
}

#[test]
fn alternate_content_handles_nested_empty_and_missing_fallback_branches() {
    let nested = format!(
        r#"<mc:AlternateContent><mc:Choice Requires="unknown">{}</mc:Choice>
        <mc:Fallback>{}</mc:Fallback></mc:AlternateContent>"#,
        marker("FF0000"),
        marker("00FF00")
    );
    let xml = format!(
        r#"<mc:AlternateContent xmlns:mc="{MC}" xmlns:a14="{A14}">
        <mc:Choice Requires="a14">{nested}</mc:Choice><mc:Fallback>{}</mc:Fallback>
        </mc:AlternateContent>
        <mc:AlternateContent xmlns:mc="{MC}" xmlns:a14="{A14}"><mc:Choice Requires="a14"/><mc:Fallback>{}</mc:Fallback></mc:AlternateContent>
        <mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="unknown">{}</mc:Choice></mc:AlternateContent>"#,
        marker("FFFFFF"),
        marker("FFFFFF"),
        marker("FFFFFF")
    );
    assert_eq!(colors(&xml), vec![Color::new(0, 255, 0)]);
}

#[test]
fn alternate_content_in_groups_keeps_ancestor_namespace_bindings() {
    let xml = format!(
        r#"<p:grpSp xmlns:a14="{A14}"><p:nvGrpSpPr/><p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>
        <a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>
        <mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="a14">{}</mc:Choice>
        <mc:Fallback>{}</mc:Fallback></mc:AlternateContent></p:grpSp>"#,
        marker("00FF00"),
        marker("FF0000")
    );
    assert_eq!(colors(&xml), vec![Color::new(0, 255, 0)]);
}

#[test]
fn alternate_content_math_fixture_keeps_native_content_without_bitmap_duplicate() {
    let data = include_bytes!("../../../../tests/fixtures/pptx/office_math_alternate_content.pptx");
    let (doc, _) = PptxParser.parse(data, &ConvertOptions::default()).unwrap();
    assert_eq!(doc.pages.len(), 2);
    for page in &doc.pages {
        let Page::Fixed(page) = page else {
            panic!("fixed slide")
        };
        assert_eq!(
            page.elements
                .iter()
                .filter(|e| matches!(e.kind, FixedElementKind::Image(_)))
                .count(),
            0
        );
        assert_eq!(
            page.elements
                .iter()
                .filter(|e| matches!(e.kind, FixedElementKind::TextBox(_)))
                .count(),
            3, // Title, explanatory caption, and one native equation.
        );
    }
}
