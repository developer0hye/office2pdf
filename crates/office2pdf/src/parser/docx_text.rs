use super::{
    Alignment, Color, HyperlinkMap, LineSpacing, ParagraphStyle, TabAlignment, TabLeader, TabStop,
    TabStopOverride, TextStyle, VerticalTextAlign, apply_tab_stop_overrides,
};
use crate::parser::units::{half_points_to_pt, twips_to_pt};
use crate::parser::xml_util;

// Word supplies an application- and locale-dependent sans face when the
// package omits one; Arial gives the parser a stable cross-platform baseline.
const WORD_COMPATIBLE_DEFAULT_FONT: &str = "Arial";

pub(super) fn extract_paragraph_style(prop: &docx_rs::ParagraphProperty) -> ParagraphStyle {
    let alignment = prop.alignment.as_ref().and_then(|j| match j.val.as_str() {
        "center" => Some(Alignment::Center),
        "right" | "end" => Some(Alignment::Right),
        "left" | "start" => Some(Alignment::Left),
        "both" | "justified" => Some(Alignment::Justify),
        _ => None,
    });

    let (indent_left, indent_right, indent_first_line) = extract_indent(&prop.indent);
    let (line_spacing, space_before, space_after) = extract_line_spacing(&prop.line_spacing);
    let tab_stops = extract_tab_stops(&prop.tabs);

    ParagraphStyle {
        alignment,
        indent_left,
        indent_right,
        indent_first_line,
        line_spacing,
        line_box: None,
        space_before,
        space_after,
        heading_level: None,
        direction: None,
        tab_stops,
    }
}

fn extract_indent(indent: &Option<docx_rs::Indent>) -> (Option<f64>, Option<f64>, Option<f64>) {
    let Some(indent) = indent else {
        return (None, None, None);
    };

    let left = indent.start.map(twips_to_pt);
    let right = indent.end.map(twips_to_pt);
    let first_line = indent.special_indent.map(|si| match si {
        docx_rs::SpecialIndentType::FirstLine(v) => twips_to_pt(v),
        docx_rs::SpecialIndentType::Hanging(v) => -twips_to_pt(v),
    });

    (left, right, first_line)
}

fn extract_line_spacing(
    spacing: &Option<docx_rs::LineSpacing>,
) -> (Option<LineSpacing>, Option<f64>, Option<f64>) {
    let Some(spacing) = spacing else {
        return (None, None, None);
    };

    let json = match serde_json::to_value(spacing) {
        Ok(j) => j,
        Err(_) => return (None, None, None),
    };

    let space_before = json.get("before").and_then(|v| v.as_f64()).map(twips_to_pt);
    let space_after = json.get("after").and_then(|v| v.as_f64()).map(twips_to_pt);

    let line_spacing = json.get("line").and_then(|line_val| {
        let line = line_val.as_f64()?;
        let rule = json.get("lineRule").and_then(|v| v.as_str());
        match rule {
            Some("exact") | Some("atLeast") => Some(LineSpacing::Exact(twips_to_pt(line))),
            _ => Some(LineSpacing::Proportional(line / 240.0)),
        }
    });

    (line_spacing, space_before, space_after)
}

pub(super) fn extract_tab_stops(tabs: &[docx_rs::Tab]) -> Option<Vec<TabStop>> {
    let tab_overrides = extract_tab_stop_overrides(tabs)?;
    let mut tab_stops: Vec<TabStop> = Vec::new();
    apply_tab_stop_overrides(&mut tab_stops, &tab_overrides);
    Some(tab_stops)
}

pub(super) fn extract_tab_stop_overrides(tabs: &[docx_rs::Tab]) -> Option<Vec<TabStopOverride>> {
    if tabs.is_empty() {
        return None;
    }

    Some(
        tabs.iter()
            .filter_map(|tab| {
                let position = tab.pos.map(|pos_twips| twips_to_pt(pos_twips as f64))?;

                if matches!(tab.val, Some(docx_rs::TabValueType::Clear)) {
                    return Some(TabStopOverride::Clear(position));
                }

                let alignment = match tab.val {
                    Some(docx_rs::TabValueType::Center) => TabAlignment::Center,
                    Some(docx_rs::TabValueType::Right) | Some(docx_rs::TabValueType::End) => {
                        TabAlignment::Right
                    }
                    Some(docx_rs::TabValueType::Decimal) => TabAlignment::Decimal,
                    _ => TabAlignment::Left,
                };

                let leader =
                    match tab.leader {
                        Some(docx_rs::TabLeaderType::Dot)
                        | Some(docx_rs::TabLeaderType::MiddleDot) => TabLeader::Dot,
                        Some(docx_rs::TabLeaderType::Hyphen)
                        | Some(docx_rs::TabLeaderType::Heavy) => TabLeader::Hyphen,
                        Some(docx_rs::TabLeaderType::Underscore) => TabLeader::Underscore,
                        _ => TabLeader::None,
                    };

                Some(TabStopOverride::Set(TabStop {
                    position,
                    alignment,
                    leader,
                }))
            })
            .collect(),
    )
}

pub(super) fn extract_run_style(rp: &docx_rs::RunProperty) -> TextStyle {
    let json = serde_json::to_value(rp).unwrap_or(serde_json::Value::Null);
    extract_run_style_from_json(&json)
}

pub(super) fn extract_run_style_from_json(rp: &serde_json::Value) -> TextStyle {
    let vertical_align: Option<VerticalTextAlign> =
        rp.get("vertAlign").and_then(|va| match va.as_str()? {
            "superscript" => Some(VerticalTextAlign::Superscript),
            "subscript" => Some(VerticalTextAlign::Subscript),
            _ => None,
        });

    let all_caps: Option<bool> = rp.get("caps").and_then(serde_json::Value::as_bool);

    TextStyle {
        bold: rp.get("bold").and_then(serde_json::Value::as_bool),
        italic: rp.get("italic").and_then(serde_json::Value::as_bool),
        underline: rp
            .get("underline")
            .and_then(|u| u.as_str())
            .and_then(|val| if val == "none" { None } else { Some(true) }),
        strikethrough: rp.get("strike").and_then(json_bool_or_val),
        font_size: rp
            .get("sz")
            .and_then(serde_json::Value::as_f64)
            .map(half_points_to_pt),
        color: rp
            .get("color")
            .and_then(serde_json::Value::as_str)
            .and_then(xml_util::parse_hex_color),
        font_family: rp.get("fonts").and_then(|fonts| {
            fonts
                .get("ascii")
                .or_else(|| fonts.get("hiAnsi"))
                .or_else(|| fonts.get("eastAsia"))
                .or_else(|| fonts.get("cs"))
                .and_then(serde_json::Value::as_str)
                .map(String::from)
        }),
        highlight: rp
            .get("highlight")
            .and_then(serde_json::Value::as_str)
            .and_then(resolve_highlight_color),
        vertical_align,
        all_caps,
        small_caps: None,
        letter_spacing: rp
            .get("characterSpacing")
            .and_then(serde_json::Value::as_i64)
            .map(|twips| twips_to_pt(twips as f64)),
    }
}

fn json_bool_or_val(value: &serde_json::Value) -> Option<bool> {
    value
        .as_bool()
        .or_else(|| value.get("val").and_then(serde_json::Value::as_bool))
}

pub(super) fn extract_doc_default_text_style_with_theme(
    styles: &docx_rs::Styles,
    theme_fonts: &ThemeFonts,
) -> TextStyle {
    let json = serde_json::to_value(&styles.doc_defaults).ok();
    let run_property = json.as_ref().and_then(|value| {
        value
            .get("runPropertyDefault")
            .and_then(|value| value.get("runProperty"))
    });
    let mut style = run_property
        .map(extract_run_style_from_json)
        .unwrap_or_default();
    if style.font_family.is_none() {
        style.font_family = run_property
            .and_then(|property| resolve_theme_font_family(property, theme_fonts))
            .or_else(|| Some(WORD_COMPATIBLE_DEFAULT_FONT.to_string()));
    }
    style
}

/// Latin typefaces of the document theme's minor (body) and major (heading)
/// font schemes, from `word/theme/theme1.xml`.
#[derive(Debug, Clone, Default)]
pub(super) struct ThemeFonts {
    pub(super) minor_latin: Option<String>,
    pub(super) major_latin: Option<String>,
}

/// Parse the theme's font scheme latin typefaces.
pub(super) fn parse_theme_fonts(theme_xml: &str) -> ThemeFonts {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    let mut fonts = ThemeFonts::default();
    let mut reader = Reader::from_str(theme_xml);
    let mut in_minor = false;
    let mut in_major = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"minorFont" => in_minor = true,
                b"majorFont" => in_major = true,
                _ => {}
            },
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"latin" => {
                let typeface: Option<String> = e.attributes().flatten().find_map(|attr| {
                    (attr.key.local_name().as_ref() == b"typeface")
                        .then(|| attr.unescape_value().ok())
                        .flatten()
                        .map(|v| v.to_string())
                        .filter(|v| !v.is_empty())
                });
                if in_minor {
                    fonts.minor_latin = typeface;
                } else if in_major {
                    fonts.major_latin = typeface;
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"minorFont" => in_minor = false,
                b"majorFont" => in_major = false,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    fonts
}

/// Resolve rFonts theme slots (asciiTheme="minorHAnsi" etc.) against the
/// document theme when no literal font family is given.
pub(super) fn resolve_theme_font_family(
    run_property_json: &serde_json::Value,
    theme_fonts: &ThemeFonts,
) -> Option<String> {
    let fonts = run_property_json.get("fonts")?;
    let slot: &str = fonts
        .get("asciiTheme")
        .or_else(|| fonts.get("hiAnsiTheme"))
        .or_else(|| fonts.get("eastAsiaTheme"))
        .or_else(|| fonts.get("csTheme"))
        .and_then(serde_json::Value::as_str)?;
    if slot.starts_with("minor") {
        theme_fonts.minor_latin.clone()
    } else if slot.starts_with("major") {
        theme_fonts.major_latin.clone()
    } else {
        None
    }
}

pub(super) fn resolve_highlight_color(name: &str) -> Option<Color> {
    match name {
        "yellow" => Some(Color::new(255, 255, 0)),
        "green" => Some(Color::new(0, 255, 0)),
        "cyan" => Some(Color::new(0, 255, 255)),
        "magenta" => Some(Color::new(255, 0, 255)),
        "blue" => Some(Color::new(0, 0, 255)),
        "red" => Some(Color::new(255, 0, 0)),
        "darkBlue" => Some(Color::new(0, 0, 128)),
        "darkCyan" => Some(Color::new(0, 128, 128)),
        "darkGreen" => Some(Color::new(0, 128, 0)),
        "darkMagenta" => Some(Color::new(128, 0, 128)),
        "darkRed" => Some(Color::new(128, 0, 0)),
        "darkYellow" => Some(Color::new(128, 128, 0)),
        "darkGray" => Some(Color::new(128, 128, 128)),
        "lightGray" => Some(Color::new(192, 192, 192)),
        "black" => Some(Color::new(0, 0, 0)),
        "white" => Some(Color::new(255, 255, 255)),
        _ => None,
    }
}

// Re-export for sibling modules that import from here.
pub(super) use xml_util::parse_hex_color;

pub(super) fn resolve_hyperlink_url(
    hyperlink: &docx_rs::Hyperlink,
    hyperlinks: &HyperlinkMap,
) -> Option<String> {
    match &hyperlink.link {
        docx_rs::HyperlinkData::External { rid, path } => {
            if !path.is_empty() {
                Some(path.clone())
            } else {
                hyperlinks.get(rid).cloned()
            }
        }
        docx_rs::HyperlinkData::Anchor { .. } => None,
    }
}

pub(super) fn is_column_break(br: &docx_rs::Break) -> bool {
    break_type(br).as_deref() == Some("column")
}

pub(super) fn is_page_break(br: &docx_rs::Break) -> bool {
    break_type(br).as_deref() == Some("page")
}

fn break_type(br: &docx_rs::Break) -> Option<String> {
    serde_json::to_value(br)
        .ok()
        .and_then(|value| value.get("breakType")?.as_str().map(String::from))
}

pub(super) fn extract_run_text_skip_layout_breaks(run: &docx_rs::Run) -> String {
    let mut text = String::new();
    for child in &run.children {
        match child {
            docx_rs::RunChild::Text(t) => text.push_str(&t.text),
            docx_rs::RunChild::Tab(_) => text.push('\t'),
            docx_rs::RunChild::Break(br) if !is_column_break(br) && !is_page_break(br) => {
                text.push('\n');
            }
            _ => {}
        }
    }
    text
}

pub(super) fn extract_run_text(run: &docx_rs::Run) -> String {
    let mut text = String::new();
    for child in &run.children {
        match child {
            docx_rs::RunChild::Text(t) => text.push_str(&t.text),
            docx_rs::RunChild::Tab(_) => text.push('\t'),
            docx_rs::RunChild::Break(_) => text.push('\n'),
            _ => {}
        }
    }
    text
}

/// Extract the referenced character style id (`<w:rStyle>`) from a run's
/// properties, if present. docx-rs serialises the reference under the `style`
/// key. Used to resolve syntax-highlighting token styles (issue #176).
pub(super) fn extract_run_style_id(run_property: &docx_rs::RunProperty) -> Option<String> {
    serde_json::to_value(run_property)
        .ok()?
        .get("style")?
        .as_str()
        .map(String::from)
}
