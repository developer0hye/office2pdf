use super::{
    Alignment, Color, HyperlinkMap, LineSpacing, ParagraphStyle, TabAlignment, TabLeader, TabStop,
    TabStopOverride, TextStyle, VerticalTextAlign, apply_tab_stop_overrides,
};

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
        grid_line_pitch: None,
        space_before,
        space_after,
        heading_level: None,
        direction: None,
        east_asian_line_break: None,
        tab_stops,
        container: None,
    }
}

fn extract_indent(indent: &Option<docx_rs::Indent>) -> (Option<f64>, Option<f64>, Option<f64>) {
    let Some(indent) = indent else {
        return (None, None, None);
    };

    let left = indent.start.map(|v| v as f64 / 20.0);
    let right = indent.end.map(|v| v as f64 / 20.0);
    let first_line = indent.special_indent.map(|si| match si {
        docx_rs::SpecialIndentType::FirstLine(v) => v as f64 / 20.0,
        docx_rs::SpecialIndentType::Hanging(v) => -(v as f64 / 20.0),
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

    let space_before = json
        .get("before")
        .and_then(|v| v.as_f64())
        .map(|v| v / 20.0);
    let space_after = json.get("after").and_then(|v| v.as_f64()).map(|v| v / 20.0);

    let line_spacing = json.get("line").and_then(|line_val| {
        let line = line_val.as_f64()?;
        let rule = json.get("lineRule").and_then(|v| v.as_str());
        match rule {
            Some("exact") | Some("atLeast") => Some(LineSpacing::Exact(line / 20.0)),
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
                let position = tab.pos.map(|pos_twips| pos_twips as f64 / 20.0)?;

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
    let (font_family_ascii, font_family_hansi, font_family_east_asia, font_family_cs): (
        Option<String>,
        Option<String>,
        Option<String>,
        Option<String>,
    ) = rp
        .get("fonts")
        .map_or((None, None, None, None), extract_font_slots);
    let font_family: Option<String> = font_family_ascii
        .clone()
        .or_else(|| font_family_hansi.clone())
        .or_else(|| font_family_east_asia.clone())
        .or_else(|| font_family_cs.clone());

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
            .map(|half_points| half_points / 2.0),
        color: rp
            .get("color")
            .and_then(serde_json::Value::as_str)
            .and_then(parse_hex_color),
        font_family,
        font_family_ascii,
        font_family_hansi,
        font_family_east_asia,
        font_family_cs,
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
            .map(|twips| twips as f64 / 20.0),
    }
}

fn extract_font_slots(
    fonts: &serde_json::Value,
) -> (
    Option<String>,
    Option<String>,
    Option<String>,
    Option<String>,
) {
    (
        extract_font_slot_value(fonts, &["ascii"]),
        extract_font_slot_value(fonts, &["hiAnsi", "hAnsi", "highAnsi"]),
        extract_font_slot_value(fonts, &["eastAsia"]),
        extract_font_slot_value(fonts, &["cs"]),
    )
}

fn extract_font_slot_value(fonts: &serde_json::Value, keys: &[&str]) -> Option<String> {
    keys.iter()
        .find_map(|key| fonts.get(*key).and_then(serde_json::Value::as_str))
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(String::from)
}

fn json_bool_or_val(value: &serde_json::Value) -> Option<bool> {
    value
        .as_bool()
        .or_else(|| value.get("val").and_then(serde_json::Value::as_bool))
}

pub(super) fn extract_doc_default_text_style(styles: &docx_rs::Styles) -> TextStyle {
    let Ok(json) = serde_json::to_value(&styles.doc_defaults) else {
        return TextStyle::default();
    };
    let Some(run_property) = json
        .get("runPropertyDefault")
        .and_then(|value| value.get("runProperty"))
    else {
        return TextStyle::default();
    };

    extract_run_style_from_json(run_property)
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

pub(super) fn parse_hex_color(hex: &str) -> Option<Color> {
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some(Color::new(r, g, b))
}

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
    serde_json::to_value(br)
        .ok()
        .and_then(|v| {
            v.get("breakType")
                .and_then(|bt| bt.as_str().map(|s| s == "column"))
        })
        .unwrap_or(false)
}

pub(super) fn extract_run_text_skip_column_breaks(run: &docx_rs::Run) -> String {
    let mut text = String::new();
    for child in &run.children {
        match child {
            docx_rs::RunChild::Text(t) => text.push_str(&t.text),
            docx_rs::RunChild::Tab(_) => text.push('\t'),
            docx_rs::RunChild::Break(br) => {
                if !is_column_break(br) {
                    text.push('\n');
                }
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
