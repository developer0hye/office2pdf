use super::{
    Alignment, Color, HyperlinkMap, LineSpacing, PairKerning, ParagraphStyle, TabAlignment,
    TabLeader, TabStop, TabStopOverride, TextStyle, VerticalTextAlign, apply_tab_stop_overrides,
};
use crate::ir::{BorderLineStyle, BorderSide, CellBorder, Insets, LineJoin, Run};
use crate::parser::units::{half_points_to_pt, twips_to_pt};
use crate::parser::xml_util;
use crate::parser::xml_util::OOXML_XML_VERSION;

// Word's fallback for a package that resolves no Latin face at all — no
// `w:rFonts` in `w:docDefaults`, none on a style, none on a run, and no theme
// slot to resolve through — is a serif, not a sans. `pdffonts` on a native
// macOS Word 16 export of `tests/fixtures/docx/unit_test_headers.docx`, which
// is exactly that package, embeds `TimesNewRomanPSMT` and
// `TimesNewRomanPS-BoldMT`; the `Arial` this named before rendered the page
// sans and set its 24pt `Heading1` line 11% wide — 417.41pt of visible ink
// against the export's 375.96pt (issue #1196).
//
// Unmeasured: whether a package that carries a `word/theme/theme1.xml` while
// still naming no `w:rFonts` takes `minorHAnsi` instead. Nothing reaches the
// theme without an attribute naming a slot — `resolve_theme_font_family`
// requires one — so both readings agree on every package that has no theme
// part, which is every fixture in the corpus of this shape.
const WORD_COMPATIBLE_DEFAULT_FONT: &str = "Times New Roman";

/// Map a `w:jc` value onto the IR's alignment.
///
/// Shared with the table-style resolver so a style's `w:jc` and a paragraph's
/// own `w:jc` cannot disagree — they did, over `distribute`, which one side
/// dropped to `None` while the other justified it (issue #845).
///
/// `distribute` is Word's East Asian distributed justification: like `both`,
/// but it stretches the last line too. The IR has one justify mode, so both
/// land on it.
pub(super) fn parse_alignment(value: &str) -> Option<Alignment> {
    match value {
        "center" => Some(Alignment::Center),
        "right" | "end" => Some(Alignment::Right),
        "left" | "start" => Some(Alignment::Left),
        "both" | "justified" | "distribute" => Some(Alignment::Justify),
        _ => None,
    }
}

pub(super) fn extract_paragraph_style(prop: &docx_rs::ParagraphProperty) -> ParagraphStyle {
    let alignment = prop
        .alignment
        .as_ref()
        .and_then(|justification| parse_alignment(justification.val.as_str()));

    let (indent_left, indent_right, indent_first_line) = extract_indent(&prop.indent);
    let (line_spacing, space_before, space_after) = extract_line_spacing(&prop.line_spacing);
    let tab_stops = extract_tab_stops(&prop.tabs);
    let border = extract_paragraph_borders(&prop.borders);
    let border_space = border
        .as_ref()
        .and_then(|_| extract_paragraph_border_space(&prop.borders));

    ParagraphStyle {
        alignment,
        // Not read from `prop`: the published docx-rs does not parse
        // `w:wordWrap`, and reading the patched fork's field made the crate
        // unpublishable (issue #1041). The raw-XML word-wrap context supplies
        // the paragraph value; styles take theirs from `scan_style_word_wrap`.
        word_wrap: None,
        indent_left,
        indent_right,
        indent_first_line,
        line_spacing,
        line_box: None,
        space_before,
        space_after,
        space_before_percent: None,
        space_after_percent: None,
        heading_level: None,
        direction: None,
        tab_stops,
        default_tab_stop_pt: None,
        background: None,
        border,
        border_space,
        // Word does not share a fixed line box across a
        // paragraph's fonts, so its paragraph mark carries no line-box role.
        paragraph_mark_font_family: None,
        // Worksheet-cell-only (issue #1631); Word paragraphs never carry one.
        sheet_number_format_reserved_glyphs: None,
    }
}

/// Each `w:pBdr` side's `w:space`, in points — the gap Word leaves between the
/// paragraph text and that rule. The attribute's own default is 0, so a border
/// that omits it gets no gap; substituting a fixed 4pt for every document
/// displaced everything below a bordered paragraph by the difference
/// (issue #520).
fn extract_paragraph_border_space(
    borders: &Option<docx_rs::ParagraphBorders>,
) -> Option<Box<Insets>> {
    let json = serde_json::to_value(borders.as_ref()?).ok()?;
    let side_space = |name: &str| -> f64 {
        json.get(name)
            .and_then(|side| side.get("space"))
            .and_then(serde_json::Value::as_f64)
            .unwrap_or(0.0)
    };
    Some(Box::new(Insets {
        top: side_space("top"),
        right: side_space("right"),
        bottom: side_space("bottom"),
        left: side_space("left"),
    }))
}

/// Word draws `w:pPr/w:pBdr` rules around the full paragraph width (heading
/// underlines, letterhead frames). docx-rs keeps the side fields private, so
/// they are read through the serialized form; `w:sz` is eighths of a point.
fn extract_paragraph_borders(
    borders: &Option<docx_rs::ParagraphBorders>,
) -> Option<Box<CellBorder>> {
    let borders = borders.as_ref()?;
    let json = serde_json::to_value(borders).ok()?;

    let side = |name: &str| -> Option<BorderSide> {
        let side_json = json.get(name)?;
        let val = side_json
            .get("val")
            .or_else(|| side_json.get("borderType"))?
            .as_str()?;
        // The thinThick/thickThin families are Word's multi-line borders, so
        // they carry the double rule rather than the single one a bare
        // `_` arm would give them. Header and footer paragraphs used to read
        // their borders through a second extractor that had this list while
        // this one did not; both now come through here (issue #1822).
        let style = match val {
            "nil" | "none" => return None,
            "double"
            | "triple"
            | "thinThickSmallGap"
            | "thickThinSmallGap"
            | "thinThickMediumGap"
            | "thickThinMediumGap"
            | "thinThickLargeGap"
            | "thickThinLargeGap"
            | "thinThickThinSmallGap"
            | "thinThickThinMediumGap"
            | "thinThickThinLargeGap" => BorderLineStyle::Double,
            "dotted" => BorderLineStyle::Dotted,
            "dashed" | "dashSmallGap" => BorderLineStyle::Dashed,
            "dotDash" | "dashDotStroked" => BorderLineStyle::DashDot,
            "dotDotDash" => BorderLineStyle::DashDotDot,
            _ => BorderLineStyle::Solid,
        };
        let size = side_json
            .get("size")
            .and_then(|v| v.as_f64())
            .unwrap_or(4.0);
        let color = side_json
            .get("color")
            .and_then(|v| v.as_str())
            .filter(|value| *value != "auto")
            .and_then(xml_util::parse_hex_color)
            .unwrap_or_else(Color::black);
        Some(BorderSide {
            width: size / 8.0,
            color,
            style,
            join: LineJoin::Round,
        })
    };

    let border = CellBorder {
        top: side("top"),
        bottom: side("bottom"),
        left: side("left"),
        right: side("right"),
    };
    if border.top.is_none()
        && border.bottom.is_none()
        && border.left.is_none()
        && border.right.is_none()
    {
        return None;
    }
    Some(Box::new(border))
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

    line_spacing_from_json(&json)
}

fn line_spacing_from_json(
    json: &serde_json::Value,
) -> (Option<LineSpacing>, Option<f64>, Option<f64>) {
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
        // Word shapes East Asian codepoints with `w:eastAsia` and Latin ones
        // with `w:ascii` in the same run. Collapsing the two into one family
        // dropped whichever came second, so Hangul was shaped by falling back
        // from the Latin family instead (issue #575).
        east_asian_font_family: rp.get("fonts").and_then(|fonts| {
            fonts
                .get("eastAsia")
                .and_then(serde_json::Value::as_str)
                .map(String::from)
        }),
        // `w:rPr/w:color` is a bare hex value with no opacity channel.
        color_alpha: None,
        highlight: rp
            .get("highlight")
            .and_then(serde_json::Value::as_str)
            .and_then(resolve_highlight_color),
        vertical_align,
        baseline_shift: None,
        all_caps,
        small_caps: None,
        letter_spacing: rp
            .get("characterSpacing")
            .and_then(serde_json::Value::as_i64)
            .map(|twips| twips_to_pt(twips as f64)),
        pair_kerning: extract_pair_kerning(rp),
    }
}

/// Read `w:rPr/w:kern` (ECMA-376 §17.3.2.15) into the IR's kerning model.
///
/// The element is a *size threshold* in half-points, not a switch: Word kerns
/// a run only once its size reaches `w:val`. `None` means the properties state
/// nothing, which at every level below `w:docDefaults` means *inherit* — only
/// `PairKerningRules` (in `docx_styles.rs`) turns an absent element into a
/// decision, and only for the document default, where Word's own answer is
/// not to kern. Leaving the OpenType feature on where Word does not kern
/// tightened display headings by up to 2.02pt at 22pt and, because they are
/// centred, shifted them right (issue #628).
///
/// docx-rs drops the element on the way in — its `RunProperty` has no field
/// for it and its JSON never carries a `kern` key — so a real file reaches
/// this function only through the raw reader's JSON-shaped tests. The document
/// default and every named style are read from `word/styles.xml` directly;
/// see `PairKerningRules` for what that covers and what it does not.
pub(super) fn extract_pair_kerning(rp: &serde_json::Value) -> Option<PairKerning> {
    let threshold_half_points: f64 = rp.get("kern").and_then(|kern| {
        kern.as_f64()
            .or_else(|| kern.get("val").and_then(serde_json::Value::as_f64))
    })?;
    Some(pair_kerning_from_half_points(threshold_half_points))
}

/// Turn a stated `w:kern w:val` into the IR's kerning model.
///
/// `w:val="0"` is how Word records "kerning off" explicitly; it is a stated
/// decision rather than "kern from 0pt up".
pub(super) fn pair_kerning_from_half_points(half_points: f64) -> PairKerning {
    PairKerning::from_threshold_pt(half_points_to_pt(half_points))
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

/// The document-wide paragraph default, `w:docDefaults/w:pPrDefault`.
///
/// It sits at the bottom of the paragraph cascade — below the `w:default="1"`
/// paragraph style, below every named style, below direct formatting — and is
/// where a generated document normally states its body justification, line
/// spacing, and space-after. Reading only `w:rPrDefault` left every paragraph
/// that relied on it ragged, single-spaced, and gapless: the technical-brief
/// fixture set a 15.87pt line advance against Word's 19.92pt and paginated to
/// 31 pages instead of 39 (issue #574).
///
/// docx-rs keeps `DocDefaults`' fields private, so this reads the serialized
/// form, as `extract_doc_default_text_style_with_theme` does above.
pub(super) fn extract_doc_default_paragraph_style(styles: &docx_rs::Styles) -> ParagraphStyle {
    let Ok(json) = serde_json::to_value(&styles.doc_defaults) else {
        return ParagraphStyle::default();
    };
    let Some(paragraph_property) = json
        .get("paragraphPropertyDefault")
        .and_then(|value| value.get("paragraphProperty"))
    else {
        return ParagraphStyle::default();
    };

    let alignment = paragraph_property
        .get("alignment")
        .and_then(serde_json::Value::as_str)
        .and_then(parse_alignment);
    let (line_spacing, space_before, space_after) = paragraph_property
        .get("lineSpacing")
        .map(line_spacing_from_json)
        .unwrap_or((None, None, None));

    ParagraphStyle {
        alignment,
        line_spacing,
        space_before,
        space_after,
        ..ParagraphStyle::default()
    }
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
                        .then(|| attr.normalized_value(OOXML_XML_VERSION).ok())
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

/// In-text marker for Word's automatic space between East Asian text and
/// adjacent Latin letters or digits. The renderer expands it to a quarter em
/// and it is never emitted literally; see the matching constant there.
const EAST_ASIAN_AUTO_SPACE_CHAR: char = '\u{E001}';

/// Insert Word's automatic space at every East Asian/Latin boundary in the
/// paragraph that does not already carry a literal space.
///
/// The space is exactly a quarter em of the run's size — 2.625pt at 10.5pt,
/// 2.375pt at 9.5pt — applied on both sides of a Latin island (`제3자` widens
/// at `제→3` and at `3→자`), measured on native Word exports (issue #521).
///
/// The caller decides *whether* a paragraph is eligible; this function only
/// knows where the boundaries are. Eligibility is a property of the
/// paragraph's *style resolution*, settled by a one-factor probe (issue
/// #732): Word's built-in Korean Normal style suppresses the space, and any
/// explicitly defined style — a resolvable `w:pStyle`, or a defined default
/// paragraph style reached by bare paragraphs — replaces the built-in and
/// restores the spec default of on. Adding only a `Normal` definition to a
/// package flips every bare-paragraph, cell and justified boundary from flush
/// to +0.25em; removing it from #521's probe flips them all back. That one
/// factor is why #521's probe (which defined `Normal`) read the space as
/// unconditional while the corpus mocks (which define none) draw their bare
/// paragraphs, cells (issue #627) and centred date line (issue #728) flush
/// yet widen their `ListParagraph` items.
///
/// It is a property of *nothing else*: a second one-factor probe patched only
/// `w:jc` in that same package and measured left, centred, justified and right
/// at the same +2.588pt per boundary (issue #1053).
///
/// The boundary can fall between two runs, so the scan carries the previous
/// character across the run break rather than restarting at each run.
///
/// TODO(docx-rs models neither `w:autoSpaceDE` nor `w:autoSpaceDN`, so a
/// paragraph that explicitly turns them off still gets the space; both default
/// to on, which is what every file in the corpus relies on).
pub(super) fn insert_east_asian_auto_space(runs: &mut [Run]) {
    let mut previous: Option<char> = None;
    for run in runs.iter_mut() {
        if run.footnote.is_some() {
            previous = None;
            continue;
        }
        let mut spaced = String::with_capacity(run.text.len());
        for ch in run.text.chars() {
            // Never trust the marker from the input: it is ours to place.
            if ch == EAST_ASIAN_AUTO_SPACE_CHAR {
                continue;
            }
            if previous.is_some_and(|previous| needs_auto_space_between(previous, ch)) {
                spaced.push(EAST_ASIAN_AUTO_SPACE_CHAR);
            }
            spaced.push(ch);
            previous = Some(ch);
        }
        run.text = spaced;
    }
}

fn needs_auto_space_between(left: char, right: char) -> bool {
    (is_east_asian_text(left) && is_western_alphanumeric(right))
        || (is_western_alphanumeric(left) && is_east_asian_text(right))
}

/// East Asian *text*, deliberately narrower than the renderer's `is_cjk_like`:
/// CJK symbols and punctuation and the fullwidth forms are already full-width
/// and Word adds nothing beside them, so widening a boundary at `、` or `．`
/// would be over-applying the rule.
fn is_east_asian_text(ch: char) -> bool {
    matches!(
        ch as u32,
        0x1100..=0x11FF   // Hangul Jamo
            | 0x3040..=0x30FF // Hiragana and Katakana
            | 0x3130..=0x318F // Hangul Compatibility Jamo
            | 0x3400..=0x4DBF // CJK Unified Ideographs Extension A
            | 0x4E00..=0x9FFF // CJK Unified Ideographs
            | 0xAC00..=0xD7AF // Hangul Syllables
            | 0xF900..=0xFAFF // CJK Compatibility Ideographs
    )
}

fn is_western_alphanumeric(ch: char) -> bool {
    ch.is_ascii_alphanumeric()
}
