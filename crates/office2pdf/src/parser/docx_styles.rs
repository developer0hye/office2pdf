use crate::parser::xml_util::OOXML_XML_VERSION;
use std::collections::{HashMap, HashSet};

use crate::ir::{Color, PairKerning, ParagraphStyle, TabStop, TextStyle};

use super::{
    ThemeFonts, extract_doc_default_paragraph_style, extract_doc_default_text_style_with_theme,
    extract_paragraph_style, extract_run_style, extract_tab_stop_overrides,
    pair_kerning_from_half_points, resolve_theme_font_family,
};

/// The `w:kern` thresholds a document states outside its runs, read from the
/// raw `word/styles.xml`.
///
/// Read from the raw part for the reason `extract_default_tab_stop_pt` reads
/// `w:defaultTabStop` from raw `word/settings.xml`: the pinned docx-rs fork
/// has no field for the element, so waiting on its parse would report every
/// document as unkerned — including the 21 tracked fixtures that do state
/// `w:kern`, 9 of them as `<w:kern w:val="2"/>` in `w:docDefaults`, which asks
/// Word to kern everything (issue #628 review).
///
/// Covers the two levels a stated threshold actually reaches text from:
/// `w:docDefaults/w:rPrDefault/w:rPr/w:kern`, and each named style's
/// `w:rPr/w:kern`.
///
/// TODO(direct run `w:kern` in `word/document.xml` is not read): docx-rs drops
/// the element from `RunProperty`, and the only way to reattach it without
/// re-implementing run parsing is a positional cursor over `<w:r>` elements,
/// the shape `SmallCapsContext` uses. That cursor cannot be trusted here: the
/// scan counts runs the conversion never consumes — `w:del` runs, which
/// `flatten_tracked_changes` drops, and text-box runs, which convert through
/// their own path — so a document mixing those with a direct `w:kern` would
/// hand the threshold to the wrong run. A document whose runs state `w:kern`
/// therefore takes its style's answer, not the run's. The only tracked fixture
/// with direct run `w:kern` states `w:val="0"`, which is what the absent
/// `w:docDefaults` element already resolves to, so nothing in the corpus
/// changes. Reading it properly needs the element parsed upstream.
#[derive(Debug, Clone)]
pub(super) struct PairKerningRules {
    /// `w:docDefaults/w:rPrDefault/w:rPr/w:kern`. Absence here is a decision,
    /// not inheritance: Word ships with font kerning off, which is why the
    /// English mocks — none of which state the element — set every glyph at
    /// its nominal advance.
    document_default: PairKerning,
    /// `w:style/w:rPr/w:kern`, keyed by `w:styleId`. A style that states
    /// nothing is absent from the map and inherits.
    by_style_id: HashMap<String, PairKerning>,
}

impl Default for PairKerningRules {
    fn default() -> Self {
        Self {
            document_default: PairKerning::Never,
            by_style_id: HashMap::new(),
        }
    }
}

impl PairKerningRules {
    pub(super) fn from_styles_xml(xml: Option<&str>) -> Self {
        let Some(xml) = xml else {
            return Self::default();
        };
        Self::scan(xml)
    }

    /// The decision every run inherits when neither its style nor its own
    /// properties state one.
    pub(super) fn document_default(&self) -> PairKerning {
        self.document_default
    }

    /// What a named style states, or `None` when it states nothing and its
    /// runs inherit the document default.
    pub(super) fn for_style(&self, style_id: &str) -> Option<PairKerning> {
        self.by_style_id.get(style_id).copied()
    }

    fn scan(xml: &str) -> Self {
        use quick_xml::events::Event;

        let mut rules = Self::default();
        let mut reader = quick_xml::Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut in_doc_defaults: bool = false;
        let mut in_run_property_default: bool = false;
        // `w:pPr` never carries a `w:rPr` in a style definition (the schema
        // gives styles `CT_PPrGeneral`, which has no run properties), but a
        // malformed part must not be allowed to plant a threshold either.
        let mut in_paragraph_property: bool = false;
        let mut current_style_id: Option<String> = None;

        loop {
            match reader.read_event() {
                Ok(Event::Start(element) | Event::Empty(element)) => {
                    match element.local_name().as_ref() {
                        b"docDefaults" => in_doc_defaults = true,
                        b"rPrDefault" => in_run_property_default = true,
                        b"pPr" => in_paragraph_property = true,
                        b"style" => {
                            current_style_id = element
                                .attributes()
                                .flatten()
                                .find(|attribute| attribute.key.local_name().as_ref() == b"styleId")
                                .and_then(|attribute| {
                                    attribute
                                        .decoded_and_normalized_value(
                                            OOXML_XML_VERSION,
                                            reader.decoder(),
                                        )
                                        .ok()
                                        .map(|value| value.into_owned())
                                });
                        }
                        b"kern" if !in_paragraph_property => {
                            let Some(half_points) = read_val_attribute(&element, reader.decoder())
                            else {
                                continue;
                            };
                            let kerning: PairKerning = pair_kerning_from_half_points(half_points);
                            if in_doc_defaults && in_run_property_default {
                                rules.document_default = kerning;
                            } else if let Some(style_id) = current_style_id.as_ref() {
                                rules.by_style_id.insert(style_id.clone(), kerning);
                            }
                        }
                        _ => {}
                    }
                }
                Ok(Event::End(element)) => match element.local_name().as_ref() {
                    b"docDefaults" => in_doc_defaults = false,
                    b"rPrDefault" => in_run_property_default = false,
                    b"pPr" => in_paragraph_property = false,
                    b"style" => current_style_id = None,
                    _ => {}
                },
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }

        rules
    }
}

/// `w:val` of an element, as the number `w:kern` states it in: half-points.
fn read_val_attribute(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::Decoder,
) -> Option<f64> {
    element
        .attributes()
        .flatten()
        .find(|attribute| attribute.key.local_name().as_ref() == b"val")
        .and_then(|attribute| {
            attribute
                .decoded_and_normalized_value(OOXML_XML_VERSION, decoder)
                .ok()
        })
        .and_then(|value| value.trim().parse::<f64>().ok())
}

/// Resolved style formatting extracted from a document style definition.
/// Contains text and paragraph formatting along with an optional heading level.
pub(super) struct ResolvedStyle {
    pub(super) text: TextStyle,
    pub(super) paragraph: ParagraphStyle,
    pub(super) paragraph_tab_overrides: Option<Vec<TabStopOverride>>,
    /// Heading level from outline_lvl (0 = Heading 1, 1 = Heading 2, ..., 5 = Heading 6).
    pub(super) heading_level: Option<usize>,
    /// Whether a heading definition in this style's hierarchy supplies its
    /// own run formatting. Word does not layer a synthesized heading weight on
    /// top of such a document-defined treatment merely because `w:b` is
    /// absent (issue #1457).
    pub(super) heading_has_document_run_formatting: bool,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub(super) enum TabStopOverride {
    Set(TabStop),
    Clear(f64),
}

/// Map from style_id → resolved formatting.
pub(super) type StyleMap = HashMap<String, ResolvedStyle>;

/// Synthetic style ID used for document-level default text properties.
pub(super) const DOC_DEFAULT_STYLE_ID: &str = "__office2pdf_doc_defaults";

pub(super) fn scan_default_paragraph_style_id(styles_xml: &str) -> Option<String> {
    use quick_xml::events::Event;

    let mut reader = quick_xml::Reader::from_str(styles_xml);
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"style" =>
            {
                let mut style_type = None;
                let mut is_default = false;
                let mut style_id = None;
                for attribute in element.attributes().flatten() {
                    let value = attribute
                        .decoded_and_normalized_value(OOXML_XML_VERSION, reader.decoder())
                        .ok()
                        .map(|value| value.into_owned());
                    match attribute.key.local_name().as_ref() {
                        b"type" => style_type = value,
                        b"default" => is_default = matches!(value.as_deref(), Some("1" | "true")),
                        b"styleId" => style_id = value,
                        _ => {}
                    }
                }
                if style_type.as_deref() == Some("paragraph") && is_default {
                    return style_id;
                }
            }
            Ok(Event::Eof) | Err(_) => return None,
            _ => {}
        }
    }
}

/// Whether the document explicitly defines its default paragraph style —
/// either a paragraph style flagged `w:default="1"` (how Word writes it) or
/// one whose id is `Normal` without the flag (how docx-rs writes it).
///
/// This is the factor that decides Word's East Asian/Latin auto space for
/// bare paragraphs: Word's own built-in Korean `Normal` suppresses the space,
/// and any explicit definition replaces that built-in, restoring the spec
/// default of on. Measured on one-factor native exports — the same bytes flip
/// every bare-paragraph, cell and justified boundary between flush and
/// +0.25em on exactly this difference (issue #732).
pub(super) fn scan_defines_default_paragraph_style(styles_xml: &str) -> bool {
    use quick_xml::events::Event;

    let mut reader = quick_xml::Reader::from_str(styles_xml);
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"style" =>
            {
                let mut style_type = None;
                let mut is_default = false;
                let mut style_id = None;
                for attribute in element.attributes().flatten() {
                    let value = attribute
                        .decoded_and_normalized_value(OOXML_XML_VERSION, reader.decoder())
                        .ok()
                        .map(|value| value.into_owned());
                    match attribute.key.local_name().as_ref() {
                        b"type" => style_type = value,
                        b"default" => is_default = matches!(value.as_deref(), Some("1" | "true")),
                        b"styleId" => style_id = value,
                        _ => {}
                    }
                }
                if style_type.as_deref() == Some("paragraph")
                    && (is_default || style_id.as_deref() == Some("Normal"))
                {
                    return true;
                }
            }
            Ok(Event::Eof) | Err(_) => return false,
            _ => {}
        }
    }
}

/// Whether `w:docDefaults` declares a `w:pPrDefault` — the document taking
/// over its own paragraph defaults.
///
/// This is the factor that decides the `w:spacing w:after` a paragraph gets
/// when neither it nor its style hierarchy states one. Without the element,
/// Word's built-in `Normal` supplies 8pt; with it, the unstated value falls to
/// the spec's zero. Measured on one-factor native exports of a package that
/// states no `w:spacing` anywhere: `<w:pPrDefault/>`, `<w:pPrDefault><w:pPr/>`,
/// and a `w:pPr` carrying only `w:before` each pull the page up by the same
/// 24pt over three paragraph gaps that an explicit `w:after="0"` does, while a
/// `Normal` style carrying its own `w:pPr` leaves the export untouched
/// (issue #1085).
///
/// Read from the raw part because docx-rs models the absent element and one
/// holding an empty `w:pPr` identically — the same reason `PairKerningRules`
/// reads `w:kern` there (issue #628).
pub(super) fn scan_declares_paragraph_property_defaults(styles_xml: &str) -> bool {
    use quick_xml::events::Event;

    let mut reader = quick_xml::Reader::from_str(styles_xml);
    reader.config_mut().trim_text(true);
    // `w:pPrDefault` is only a document default inside `w:docDefaults`; the
    // scope guard keeps a stray element elsewhere in a malformed part from
    // silencing the built-in gap.
    let mut in_doc_defaults: bool = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"docDefaults" => {
                in_doc_defaults = true;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"docDefaults" => {
                in_doc_defaults = false;
            }
            Ok(Event::Start(element) | Event::Empty(element))
                if in_doc_defaults && element.local_name().as_ref() == b"pPrDefault" =>
            {
                return true;
            }
            Ok(Event::Eof) | Err(_) => return false,
            _ => {}
        }
    }
}

use crate::defaults::HEADING_FONT_SIZES;

/// Build a map from style ID → resolved formatting by extracting formatting
/// from each style's run_property and paragraph_property.
pub(super) fn build_style_map(
    styles: &docx_rs::Styles,
    theme_fonts: &ThemeFonts,
    default_paragraph_style_id: Option<&str>,
    paragraph_backgrounds: &HashMap<String, Color>,
    style_word_wraps: &HashMap<String, bool>,
    pair_kerning: &PairKerningRules,
) -> StyleMap {
    let mut map = StyleMap::new();
    let default_text: TextStyle = resolve_doc_default_text_style(styles, theme_fonts, pair_kerning);
    let default_paragraph: ParagraphStyle = extract_doc_default_paragraph_style(styles);

    map.insert(
        DOC_DEFAULT_STYLE_ID.to_string(),
        ResolvedStyle {
            text: default_text.clone(),
            paragraph: default_paragraph.clone(),
            paragraph_tab_overrides: None,
            heading_level: None,
            heading_has_document_run_formatting: false,
        },
    );

    let styles_by_id: HashMap<&str, &docx_rs::Style> = styles
        .styles
        .iter()
        .map(|style| (style.style_id.as_str(), style))
        .collect();
    for style in &styles.styles {
        if !matches!(
            style.style_type,
            docx_rs::StyleType::Paragraph | docx_rs::StyleType::Character
        ) {
            continue;
        }

        let chain = based_on_chain(style, &styles_by_id);
        let is_paragraph = style.style_type == docx_rs::StyleType::Paragraph;
        let mut resolved = ResolvedStyle {
            // Paragraph styles start at document defaults. Character styles
            // deliberately do not: applying a character style must overlay
            // only what that style hierarchy states (issue #176).
            text: if is_paragraph {
                default_text.clone()
            } else {
                TextStyle::default()
            },
            paragraph: if is_paragraph {
                default_paragraph.clone()
            } else {
                ParagraphStyle::default()
            },
            paragraph_tab_overrides: None,
            heading_level: None,
            heading_has_document_run_formatting: false,
        };

        for definition in chain.into_iter().rev() {
            merge_style_definition(
                &mut resolved,
                definition,
                theme_fonts,
                paragraph_backgrounds,
                style_word_wraps,
                pair_kerning,
            );
        }
        map.insert(style.style_id.clone(), resolved);
    }

    // Paragraphs without an explicit pStyle inherit the default paragraph
    // style (w:default="1", normally "Normal"), not just the bare document
    // defaults — fold it into the synthetic doc-default entry so its spacing,
    // line spacing, and text properties survive the cascade (issue #288).
    if let Some(style_id) = default_paragraph_style_id
        && let Some(default_style) = map.get(style_id)
    {
        let merged = ResolvedStyle {
            text: default_style.text.clone(),
            paragraph: default_style.paragraph.clone(),
            paragraph_tab_overrides: default_style.paragraph_tab_overrides.clone(),
            heading_level: None,
            heading_has_document_run_formatting: false,
        };
        map.insert(DOC_DEFAULT_STYLE_ID.to_string(), merged);
    }

    map
}

/// The style and its same-type ancestors, from child to root.
///
/// A cycle invalidates the whole parent chain for this style. Dropping the
/// chain is deterministic regardless of declaration order and safer than
/// partially inheriting a different subset depending on which cycle member
/// happened to be resolved first (issue #1453).
fn based_on_chain<'a>(
    style: &'a docx_rs::Style,
    styles_by_id: &HashMap<&str, &'a docx_rs::Style>,
) -> Vec<&'a docx_rs::Style> {
    let mut chain = vec![style];
    let mut seen = HashSet::from([style.style_id.as_str()]);
    let mut current = style;
    while let Some(parent_id) = based_on_id(current) {
        let Some(parent) = styles_by_id.get(parent_id.as_str()).copied() else {
            break;
        };
        if parent.style_type != style.style_type {
            break;
        }
        if !seen.insert(parent.style_id.as_str()) {
            return vec![style];
        }
        chain.push(parent);
        current = parent;
    }
    chain
}

fn based_on_id(style: &docx_rs::Style) -> Option<String> {
    style
        .based_on
        .as_ref()
        .and_then(|based_on| serde_json::to_value(based_on).ok())
        .and_then(|value| value.as_str().map(str::to_string))
}

fn merge_style_definition(
    resolved: &mut ResolvedStyle,
    style: &docx_rs::Style,
    theme_fonts: &ThemeFonts,
    paragraph_backgrounds: &HashMap<String, Color>,
    style_word_wraps: &HashMap<String, bool>,
    pair_kerning: &PairKerningRules,
) {
    let mut own_text = extract_run_style(&style.run_property);
    if own_text.font_family.is_none()
        && let Ok(run_property_json) = serde_json::to_value(&style.run_property)
    {
        own_text.font_family = resolve_theme_font_family(&run_property_json, theme_fonts);
    }
    // Capture provenance before adding the raw pair-kerning overlay below.
    // A kerning rule is metadata read from the same rPr, but it must not by
    // itself make an otherwise unformatted heading definition look like a complete
    // document-defined text treatment.
    let has_document_run_formatting = TextStyle {
        pair_kerning: None,
        ..own_text.clone()
    } != TextStyle::default();
    // `None` means "states nothing", so the parent decision survives.
    own_text.pair_kerning = pair_kerning.for_style(&style.style_id);
    resolved.text.merge_from(&own_text);

    if style.style_type != docx_rs::StyleType::Paragraph {
        return;
    }

    let mut own_paragraph = extract_paragraph_style(&style.paragraph_property);
    own_paragraph.background = paragraph_backgrounds.get(&style.style_id).copied();
    // From raw styles.xml because published docx-rs does not parse the field
    // (issue #1041).
    own_paragraph.word_wrap = style_word_wraps.get(&style.style_id).copied();
    let own_tab_overrides = extract_tab_stop_overrides(&style.paragraph_property.tabs);
    resolved.paragraph =
        merge_paragraph_style(&own_paragraph, own_tab_overrides.as_deref(), Some(resolved));
    // The merge above resolves clears and replacements against the inherited
    // list, so descendants consume one concrete list rather than replaying a
    // parent's raw operations.
    resolved.paragraph_tab_overrides = None;
    if let Some(level) = style
        .paragraph_property
        .outline_lvl
        .as_ref()
        .map(|outline_level| outline_level.v)
        .filter(|&value| value < 6)
    {
        resolved.heading_level = Some(level);
    }
    // Once a definition establishes the heading level, formatting on that
    // definition or any derived style is part of the document's heading
    // treatment. Formatting inherited from Normal before the heading begins
    // does not silence the bare-heading fallback (issue #1457).
    if resolved.heading_level.is_some() {
        resolved.heading_has_document_run_formatting |= has_document_run_formatting;
    }
}

/// The document-wide run defaults, with the kerning threshold the raw
/// `word/styles.xml` states folded in.
///
/// Kept apart from `extract_doc_default_text_style_with_theme` because that
/// function reads docx-rs' parse, which has no `w:kern` to give.
pub(super) fn resolve_doc_default_text_style(
    styles: &docx_rs::Styles,
    theme_fonts: &ThemeFonts,
    pair_kerning: &PairKerningRules,
) -> TextStyle {
    let mut text: TextStyle = extract_doc_default_text_style_with_theme(styles, theme_fonts);
    text.pair_kerning = Some(pair_kerning.document_default());
    text
}

/// Merge style text formatting with explicit run formatting.
/// Explicit formatting (from the run itself) takes priority over style formatting.
/// For heading styles, default sizes are applied when neither the style nor the
/// run specifies them. Bold is synthesized only for a bare, unformatted
/// heading definition; a document-defined heading run treatment that omits
/// `w:b` remains regular.
pub(super) fn merge_text_style(explicit: &TextStyle, style: Option<&ResolvedStyle>) -> TextStyle {
    let (style_text, heading_level, heading_has_document_run_formatting) = match style {
        Some(style) => (
            &style.text,
            style.heading_level,
            style.heading_has_document_run_formatting,
        ),
        None => return explicit.clone(),
    };

    let mut merged: TextStyle = style_text.clone();

    // Heading defaults: apply fallback size/bold when the style itself
    // doesn't specify them. This must happen before the explicit overwrite
    // so that explicit values still win.
    if let Some(level) = heading_level {
        if merged.font_size.is_none() {
            merged.font_size = Some(HEADING_FONT_SIZES[level]);
        }
        if merged.bold.is_none() && !heading_has_document_run_formatting {
            merged.bold = Some(true);
        }
    }

    merged.merge_from(explicit);

    merged
}

/// Merge style paragraph formatting with explicit paragraph formatting.
/// Explicit formatting takes priority.
pub(super) fn merge_paragraph_style(
    explicit: &ParagraphStyle,
    explicit_tab_overrides: Option<&[TabStopOverride]>,
    style: Option<&ResolvedStyle>,
) -> ParagraphStyle {
    let style_paragraph = style.map(|resolved_style| &resolved_style.paragraph);
    let inherited_tab_stops = style.and_then(resolve_style_tab_stops);

    ParagraphStyle {
        alignment: explicit
            .alignment
            .or(style_paragraph.and_then(|style| style.alignment)),
        // Measured on Word: a paragraph's own w:wordWrap beats the one its
        // style carries — a ListParagraph with w:val="0" breaks mid-eojeol
        // although the style alone would not (issue #730).
        word_wrap: explicit
            .word_wrap
            .or(style_paragraph.and_then(|style| style.word_wrap)),
        indent_left: explicit
            .indent_left
            .or(style_paragraph.and_then(|style| style.indent_left)),
        indent_right: explicit
            .indent_right
            .or(style_paragraph.and_then(|style| style.indent_right)),
        indent_first_line: explicit
            .indent_first_line
            .or(style_paragraph.and_then(|style| style.indent_first_line)),
        line_spacing: explicit
            .line_spacing
            .or(style_paragraph.and_then(|style| style.line_spacing)),
        line_box: explicit
            .line_box
            .or(style_paragraph.and_then(|style| style.line_box)),
        space_before: explicit
            .space_before
            .or(style_paragraph.and_then(|style| style.space_before)),
        space_after: explicit
            .space_after
            .or(style_paragraph.and_then(|style| style.space_after)),
        // Percentage paragraph gaps are a DrawingML intermediate. Word's
        // spacing reaches the IR as absolute points.
        space_before_percent: None,
        space_after_percent: None,
        heading_level: style
            .and_then(|resolved_style| resolved_style.heading_level)
            .map(|level| (level + 1) as u8),
        direction: explicit.direction,
        tab_stops: merge_tab_stops(
            explicit.tab_stops.as_deref(),
            explicit_tab_overrides,
            inherited_tab_stops.as_deref(),
        ),
        default_tab_stop_pt: explicit
            .default_tab_stop_pt
            .or(style_paragraph.and_then(|style| style.default_tab_stop_pt)),
        background: explicit
            .background
            .or(style_paragraph.and_then(|style| style.background)),
        border: explicit
            .border
            .clone()
            .or_else(|| style_paragraph.and_then(|style| style.border.clone())),
        // Follows the border it measures from rather than merging separately:
        // a paragraph that inherits its rules inherits their gaps with them,
        // and one that overrides them overrides the gaps too (issue #520).
        border_space: if explicit.border.is_some() {
            explicit.border_space.clone()
        } else {
            style_paragraph.and_then(|style| style.border_space.clone())
        },
        // Word does not share a fixed line box across a
        // paragraph's fonts, so its paragraph mark carries no line-box role.
        paragraph_mark_font_family: None,
        // Worksheet-cell-only (issue #1631); Word paragraphs never carry one.
        sheet_number_format_reserved_glyphs: None,
    }
}

fn resolve_style_tab_stops(style: &ResolvedStyle) -> Option<Vec<TabStop>> {
    resolve_tab_stop_source(
        style.paragraph.tab_stops.as_deref(),
        style.paragraph_tab_overrides.as_deref(),
    )
}

fn resolve_tab_stop_source(
    tab_stops: Option<&[TabStop]>,
    tab_overrides: Option<&[TabStopOverride]>,
) -> Option<Vec<TabStop>> {
    if let Some(tab_overrides) = tab_overrides {
        let mut resolved: Vec<TabStop> = Vec::new();
        apply_tab_stop_overrides(&mut resolved, tab_overrides);
        return Some(resolved);
    }

    tab_stops.map(|tab_stops| tab_stops.to_vec())
}

fn merge_tab_stops(
    explicit_tab_stops: Option<&[TabStop]>,
    explicit_tab_overrides: Option<&[TabStopOverride]>,
    inherited_tab_stops: Option<&[TabStop]>,
) -> Option<Vec<TabStop>> {
    if let Some(explicit_tab_overrides) = explicit_tab_overrides {
        let mut resolved: Vec<TabStop> = inherited_tab_stops.unwrap_or(&[]).to_vec();
        apply_tab_stop_overrides(&mut resolved, explicit_tab_overrides);
        return Some(resolved);
    }

    explicit_tab_stops
        .map(|tab_stops| tab_stops.to_vec())
        .or_else(|| inherited_tab_stops.map(|tab_stops| tab_stops.to_vec()))
}

pub(super) fn apply_tab_stop_overrides(
    tab_stops: &mut Vec<TabStop>,
    tab_overrides: &[TabStopOverride],
) {
    for tab_override in tab_overrides {
        match tab_override {
            TabStopOverride::Set(tab_stop) => {
                tab_stops.retain(|existing| {
                    !tab_stop_positions_match(existing.position, tab_stop.position)
                });
                tab_stops.push(*tab_stop);
            }
            TabStopOverride::Clear(position) => {
                tab_stops
                    .retain(|existing| !tab_stop_positions_match(existing.position, *position));
            }
        }
    }

    tab_stops.sort_by(|left, right| {
        left.position
            .partial_cmp(&right.position)
            .unwrap_or(std::cmp::Ordering::Equal)
    });
}

fn tab_stop_positions_match(left: f64, right: f64) -> bool {
    (left - right).abs() < 0.01
}

/// Look up the pStyle reference from a paragraph's property.
pub(super) fn get_paragraph_style_id(prop: &docx_rs::ParagraphProperty) -> Option<&str> {
    prop.style.as_ref().map(|style| style.val.as_str())
}
