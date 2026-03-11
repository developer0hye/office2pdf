use std::collections::{HashMap, HashSet};

use crate::ir::{ParagraphStyle, TabStop, TextStyle};

use super::{
    extract_doc_default_paragraph_style, extract_doc_default_text_style, extract_paragraph_style,
    extract_run_style, extract_tab_stop_overrides,
};

/// Resolved style formatting extracted from a document style definition.
/// Contains text and paragraph formatting along with an optional heading level.
#[derive(Clone)]
pub(super) struct ResolvedStyle {
    pub(super) text: TextStyle,
    pub(super) paragraph: ParagraphStyle,
    pub(super) paragraph_tab_overrides: Option<Vec<TabStopOverride>>,
    /// Heading level from outline_lvl (0 = Heading 1, 1 = Heading 2, ..., 5 = Heading 6).
    pub(super) heading_level: Option<usize>,
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

pub(super) fn normalize_style_id(style_id: &str) -> String {
    style_id.to_ascii_lowercase()
}

/// Default font sizes for heading levels (Heading 1-6).
/// Index 0 = Heading 1 (outline_lvl 0), index 5 = Heading 6 (outline_lvl 5).
const HEADING_DEFAULT_SIZES: [f64; 6] = [24.0, 20.0, 16.0, 14.0, 12.0, 11.0];

/// Build a map from style ID → resolved formatting by extracting formatting
/// from each style's run_property and paragraph_property.
pub(super) fn build_style_map(styles: &docx_rs::Styles) -> StyleMap {
    let mut map = StyleMap::new();
    let default_text: TextStyle = extract_doc_default_text_style(styles);
    let default_paragraph: ParagraphStyle = extract_doc_default_paragraph_style(styles);

    map.insert(
        DOC_DEFAULT_STYLE_ID.to_string(),
        ResolvedStyle {
            text: default_text.clone(),
            paragraph: default_paragraph.clone(),
            paragraph_tab_overrides: None,
            heading_level: None,
        },
    );

    let paragraph_styles: HashMap<String, &docx_rs::Style> = styles
        .styles
        .iter()
        .filter(|style| style.style_type == docx_rs::StyleType::Paragraph)
        .map(|style| (normalize_style_id(&style.style_id), style))
        .collect();

    let style_ids: Vec<String> = paragraph_styles.keys().cloned().collect();
    let mut resolving: HashSet<String> = HashSet::new();

    for style_id in style_ids {
        let _ = resolve_style(
            &style_id,
            &paragraph_styles,
            &default_text,
            &default_paragraph,
            &mut map,
            &mut resolving,
        );
    }

    map
}

fn resolve_style(
    style_id: &str,
    paragraph_styles: &HashMap<String, &docx_rs::Style>,
    default_text: &TextStyle,
    default_paragraph: &ParagraphStyle,
    cache: &mut StyleMap,
    resolving: &mut HashSet<String>,
) -> Option<ResolvedStyle> {
    if let Some(resolved) = cache.get(style_id) {
        return Some(resolved.clone());
    }

    let style = *paragraph_styles.get(style_id)?;
    if !resolving.insert(style_id.to_string()) {
        // Style cycles are invalid OOXML. Fall back to local style fields.
        return Some(build_resolved_style(
            style,
            None,
            default_text,
            default_paragraph,
        ));
    }

    let parent = style_based_on_id(style).and_then(|parent_id| {
        resolve_style(
            &parent_id,
            paragraph_styles,
            default_text,
            default_paragraph,
            cache,
            resolving,
        )
    });

    let resolved = build_resolved_style(style, parent.as_ref(), default_text, default_paragraph);
    resolving.remove(style_id);
    cache.insert(style_id.to_string(), resolved.clone());
    Some(resolved)
}

fn build_resolved_style(
    style: &docx_rs::Style,
    parent: Option<&ResolvedStyle>,
    default_text: &TextStyle,
    default_paragraph: &ParagraphStyle,
) -> ResolvedStyle {
    let own_text: TextStyle = extract_run_style(&style.run_property);
    let own_paragraph: ParagraphStyle = extract_paragraph_style(&style.paragraph_property);
    let own_tab_overrides: Option<Vec<TabStopOverride>> =
        extract_tab_stop_overrides(&style.paragraph_property.tabs);

    let inherited_text: &TextStyle = parent.map_or(default_text, |resolved| &resolved.text);
    let text: TextStyle = inherit_text_style(inherited_text, &own_text);

    let inherited_paragraph: Option<&ParagraphStyle> = parent
        .map(|resolved| &resolved.paragraph)
        .or(Some(default_paragraph));
    let mut paragraph: ParagraphStyle =
        inherit_paragraph_style(inherited_paragraph, &own_paragraph);

    if let Some(tab_overrides) = own_tab_overrides.as_deref() {
        let mut resolved_tabs: Vec<TabStop> = inherited_paragraph
            .and_then(|base| base.tab_stops.as_deref())
            .map_or_else(Vec::new, |tabs| tabs.to_vec());
        apply_tab_stop_overrides(&mut resolved_tabs, tab_overrides);
        paragraph.tab_stops = (!resolved_tabs.is_empty()).then_some(resolved_tabs);
    }

    let heading_level =
        extract_heading_level(style).or_else(|| parent.and_then(|resolved| resolved.heading_level));

    ResolvedStyle {
        text,
        paragraph,
        // Style-level tab inheritance is already resolved into paragraph.tab_stops.
        paragraph_tab_overrides: None,
        heading_level,
    }
}

fn style_based_on_id(style: &docx_rs::Style) -> Option<String> {
    let based_on = style.based_on.as_ref()?;
    serde_json::to_value(based_on)
        .ok()?
        .as_str()
        .map(normalize_style_id)
}

fn extract_heading_level(style: &docx_rs::Style) -> Option<usize> {
    style
        .paragraph_property
        .outline_lvl
        .as_ref()
        .map(|outline_level| outline_level.v)
        .filter(|&value| value < 6)
}

fn inherit_text_style(base: &TextStyle, override_style: &TextStyle) -> TextStyle {
    TextStyle {
        bold: override_style.bold.or(base.bold),
        italic: override_style.italic.or(base.italic),
        underline: override_style.underline.or(base.underline),
        strikethrough: override_style.strikethrough.or(base.strikethrough),
        font_size: override_style.font_size.or(base.font_size),
        color: override_style.color.or(base.color),
        font_family: override_style
            .font_family
            .clone()
            .or_else(|| base.font_family.clone()),
        font_family_ascii: override_style
            .font_family_ascii
            .clone()
            .or_else(|| base.font_family_ascii.clone()),
        font_family_hansi: override_style
            .font_family_hansi
            .clone()
            .or_else(|| base.font_family_hansi.clone()),
        font_family_east_asia: override_style
            .font_family_east_asia
            .clone()
            .or_else(|| base.font_family_east_asia.clone()),
        font_family_cs: override_style
            .font_family_cs
            .clone()
            .or_else(|| base.font_family_cs.clone()),
        highlight: override_style.highlight.or(base.highlight),
        vertical_align: override_style.vertical_align.or(base.vertical_align),
        all_caps: override_style.all_caps.or(base.all_caps),
        small_caps: override_style.small_caps.or(base.small_caps),
        letter_spacing: override_style.letter_spacing.or(base.letter_spacing),
    }
}

fn inherit_paragraph_style(
    base: Option<&ParagraphStyle>,
    override_style: &ParagraphStyle,
) -> ParagraphStyle {
    ParagraphStyle {
        alignment: override_style
            .alignment
            .or(base.and_then(|style| style.alignment)),
        indent_left: override_style
            .indent_left
            .or(base.and_then(|style| style.indent_left)),
        indent_right: override_style
            .indent_right
            .or(base.and_then(|style| style.indent_right)),
        indent_first_line: override_style
            .indent_first_line
            .or(base.and_then(|style| style.indent_first_line)),
        line_spacing: override_style
            .line_spacing
            .or(base.and_then(|style| style.line_spacing)),
        space_before: override_style
            .space_before
            .or(base.and_then(|style| style.space_before)),
        space_after: override_style
            .space_after
            .or(base.and_then(|style| style.space_after)),
        heading_level: override_style
            .heading_level
            .or(base.and_then(|style| style.heading_level)),
        direction: override_style
            .direction
            .or(base.and_then(|style| style.direction)),
        tab_stops: override_style
            .tab_stops
            .clone()
            .or_else(|| base.and_then(|style| style.tab_stops.clone())),
    }
}

/// Merge style text formatting with explicit run formatting.
/// Explicit formatting (from the run itself) takes priority over style formatting.
/// For heading styles, default sizes and bold are applied when neither the style
/// nor the run specifies them.
pub(super) fn merge_text_style(explicit: &TextStyle, style: Option<&ResolvedStyle>) -> TextStyle {
    let (style_text, heading_level) = match style {
        Some(style) => (&style.text, style.heading_level),
        None => return explicit.clone(),
    };

    let mut merged = TextStyle {
        bold: style_text.bold,
        italic: style_text.italic,
        underline: style_text.underline,
        strikethrough: style_text.strikethrough,
        font_size: style_text.font_size,
        color: style_text.color,
        font_family: style_text.font_family.clone(),
        font_family_ascii: style_text.font_family_ascii.clone(),
        font_family_hansi: style_text.font_family_hansi.clone(),
        font_family_east_asia: style_text.font_family_east_asia.clone(),
        font_family_cs: style_text.font_family_cs.clone(),
        highlight: style_text.highlight,
        vertical_align: style_text.vertical_align,
        all_caps: style_text.all_caps,
        small_caps: style_text.small_caps,
        letter_spacing: style_text.letter_spacing,
    };

    if let Some(level) = heading_level {
        if merged.font_size.is_none() {
            merged.font_size = Some(HEADING_DEFAULT_SIZES[level]);
        }
        if merged.bold.is_none() {
            merged.bold = Some(true);
        }
    }

    if explicit.bold.is_some() {
        merged.bold = explicit.bold;
    }
    if explicit.italic.is_some() {
        merged.italic = explicit.italic;
    }
    if explicit.underline.is_some() {
        merged.underline = explicit.underline;
    }
    if explicit.strikethrough.is_some() {
        merged.strikethrough = explicit.strikethrough;
    }
    if explicit.font_size.is_some() {
        merged.font_size = explicit.font_size;
    }
    if explicit.color.is_some() {
        merged.color = explicit.color;
    }
    if explicit.font_family.is_some() {
        merged.font_family = explicit.font_family.clone();
    }
    if explicit.font_family_ascii.is_some() {
        merged.font_family_ascii = explicit.font_family_ascii.clone();
    }
    if explicit.font_family_hansi.is_some() {
        merged.font_family_hansi = explicit.font_family_hansi.clone();
    }
    if explicit.font_family_east_asia.is_some() {
        merged.font_family_east_asia = explicit.font_family_east_asia.clone();
    }
    if explicit.font_family_cs.is_some() {
        merged.font_family_cs = explicit.font_family_cs.clone();
    }
    if explicit.highlight.is_some() {
        merged.highlight = explicit.highlight;
    }
    if explicit.vertical_align.is_some() {
        merged.vertical_align = explicit.vertical_align;
    }
    if explicit.all_caps.is_some() {
        merged.all_caps = explicit.all_caps;
    }
    if explicit.small_caps.is_some() {
        merged.small_caps = explicit.small_caps;
    }
    if explicit.letter_spacing.is_some() {
        merged.letter_spacing = explicit.letter_spacing;
    }

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
        space_before: explicit
            .space_before
            .or(style_paragraph.and_then(|style| style.space_before)),
        space_after: explicit
            .space_after
            .or(style_paragraph.and_then(|style| style.space_after)),
        heading_level: style
            .and_then(|resolved_style| resolved_style.heading_level)
            .map(|level| (level + 1) as u8),
        direction: explicit.direction,
        tab_stops: merge_tab_stops(
            explicit.tab_stops.as_deref(),
            explicit_tab_overrides,
            inherited_tab_stops.as_deref(),
        ),
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
