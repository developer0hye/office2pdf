use super::*;

#[test]
fn parse_hex_color_valid_red() {
    let color = parse_hex_color("FF0000").unwrap();
    assert_eq!(color, Color::new(255, 0, 0));
}

#[test]
fn parse_hex_color_valid_mixed() {
    let color = parse_hex_color("1A2B3C").unwrap();
    assert_eq!(color, Color::new(0x1A, 0x2B, 0x3C));
}

#[test]
fn parse_hex_color_wrong_length() {
    assert!(parse_hex_color("FFF").is_none());
    assert!(parse_hex_color("FF00FF00").is_none());
}

#[test]
fn parse_argb_color_valid() {
    let color = parse_argb_color("FFFF0000").unwrap();
    assert_eq!(color, Color::new(255, 0, 0));
}

#[test]
fn parse_argb_color_too_short() {
    assert!(parse_argb_color("FF00").is_none());
}

#[test]
fn resolve_relative_path_parent_traversal() {
    let result = resolve_relative_path("xl/worksheets", "../drawings/drawing1.xml");
    assert_eq!(result, "xl/drawings/drawing1.xml");
}

#[test]
fn resolve_relative_path_absolute() {
    let result = resolve_relative_path("xl/worksheets", "/xl/charts/chart1.xml");
    assert_eq!(result, "xl/charts/chart1.xml");
}

#[test]
fn resolve_relative_path_same_dir() {
    let result = resolve_relative_path("ppt/slides", "slide1.xml");
    assert_eq!(result, "ppt/slides/slide1.xml");
}

#[test]
fn resolve_relative_path_empty_base() {
    let result = resolve_relative_path("", "foo/bar.xml");
    assert_eq!(result, "foo/bar.xml");
}

#[test]
fn parse_rels_id_target_basic() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Target="slides/slide1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"/>
  <Relationship Id="rId2" Target="theme/theme1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"/>
</Relationships>"#;
    let map = parse_rels_id_target(xml);
    assert_eq!(map.get("rId1").unwrap(), "slides/slide1.xml");
    assert_eq!(map.get("rId2").unwrap(), "theme/theme1.xml");
    assert_eq!(map.len(), 2);
}

#[test]
fn skip_element_skips_nested_content() {
    let xml = r#"<root><child><nested>text</nested></child><after/></root>"#;
    let mut reader = quick_xml::Reader::from_str(xml);
    // Read <root>
    let _ = reader.read_event();
    // Read <child>
    let _ = reader.read_event();
    // Now skip child (we're inside it, depth=1)
    skip_element(&mut reader, b"child");
    // Next event should be <after/> (Empty)
    match reader.read_event() {
        Ok(quick_xml::events::Event::Empty(ref e)) => {
            assert_eq!(e.local_name().as_ref(), b"after");
        }
        other => panic!("expected Empty(after), got {other:?}"),
    }
}

#[test]
fn get_attr_str_by_qualified_name() {
    let xml = r#"<blip r:embed="rId1" xmlns:r="http://example.com"/>"#;
    let mut reader = quick_xml::Reader::from_str(xml);
    match reader.read_event() {
        Ok(quick_xml::events::Event::Empty(ref e)) => {
            assert_eq!(get_attr_str(e, b"r:embed"), Some("rId1".to_string()));
        }
        _ => panic!("expected Empty event"),
    }
}

#[test]
fn get_attr_str_by_local_name() {
    let xml = r#"<off x="100" y="200"/>"#;
    let mut reader = quick_xml::Reader::from_str(xml);
    match reader.read_event() {
        Ok(quick_xml::events::Event::Empty(ref e)) => {
            assert_eq!(get_attr_str(e, b"x"), Some("100".to_string()));
            assert_eq!(get_attr_str(e, b"y"), Some("200".to_string()));
            assert_eq!(get_attr_str(e, b"z"), None);
        }
        _ => panic!("expected Empty event"),
    }
}

#[test]
fn get_attr_i64_parses_integer() {
    let xml = r#"<off x="12700" y="-5000"/>"#;
    let mut reader = quick_xml::Reader::from_str(xml);
    match reader.read_event() {
        Ok(quick_xml::events::Event::Empty(ref e)) => {
            assert_eq!(get_attr_i64(e, b"x"), Some(12700));
            assert_eq!(get_attr_i64(e, b"y"), Some(-5000));
            assert_eq!(get_attr_i64(e, b"z"), None);
        }
        _ => panic!("expected Empty event"),
    }
}
