use std::collections::{BTreeMap, HashMap};

use crate::ir::{
    Block, List, ListItem, ListKind, ListLevelStyle, ListMarkerSuffix, NumberingPattern, Paragraph,
    ParagraphStyle, Run, TabAlignment, TabLeader, TabStop, TextStyle,
};

/// Numbering info extracted from a paragraph's numPr.
#[derive(Debug, Clone)]
pub(super) struct NumInfo {
    pub(super) num_id: usize,
    pub(super) level: u32,
    /// Under a direct `w:numPr`, which indents the paragraph's style chain
    /// supplies where its own `w:ind` is silent.
    pub(super) style_supplied_indents: StyleSuppliedIndents,
}

/// Indents a numbered paragraph takes from its style chain rather than from
/// its own `w:ind`.
#[derive(Debug, Clone, Copy, Default)]
pub(super) struct StyleSuppliedIndents {
    pub(super) left: bool,
    pub(super) first_line: bool,
}

#[derive(Debug, Clone)]
struct ResolvedListLevel {
    style: ListLevelStyle,
    paragraph_style: ParagraphStyle,
    start: u32,
    has_start_override: bool,
    restart_trigger: Option<u32>,
    /// Raw Word `lvlText` (e.g. "제%1조") for inline-numbered paragraphs.
    level_text: String,
    number_format: String,
    uses_legal_numbering: bool,
}

#[derive(Debug, Clone)]
pub(super) struct ResolvedNumbering {
    abstract_num_id: usize,
    levels: BTreeMap<u32, ResolvedListLevel>,
}

#[derive(Debug, Clone)]
struct RawListLevel {
    start: u32,
    number_format: String,
    level_text: String,
    paragraph_style: ParagraphStyle,
    marker_style: Option<TextStyle>,
    has_start_override: bool,
    uses_legal_numbering: bool,
    restart_trigger: Option<u32>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum NumberingSeries {
    Abstract(usize),
    Numbering(usize),
}

#[derive(Clone, Copy)]
struct NumberingCounter {
    value: u32,
    is_implicit: bool,
}

pub(super) type NumberingMap = HashMap<usize, ResolvedNumbering>;

fn serialize_string<T: serde::Serialize>(value: &T) -> Option<String> {
    serde_json::to_value(value)
        .ok()?
        .as_str()
        .map(|text| text.to_string())
}

fn serialize_u32<T: serde::Serialize>(value: &T) -> Option<u32> {
    serde_json::to_value(value)
        .ok()?
        .as_u64()
        .and_then(|value| u32::try_from(value).ok())
}

fn level_kind(number_format: &str) -> ListKind {
    if number_format == "bullet" {
        ListKind::Unordered
    } else {
        ListKind::Ordered
    }
}

fn counter_symbol(number_format: &str) -> Option<char> {
    match number_format {
        "decimal" | "decimalZero" => Some('1'),
        "lowerLetter" => Some('a'),
        "upperLetter" => Some('A'),
        "lowerRoman" => Some('i'),
        "upperRoman" => Some('I'),
        _ => None,
    }
}

/// Keeps lvlText's literal text apart from its counters; the renderer
/// decides whether Typst can read them back from one pattern string.
fn build_numbering_pattern(
    level_text: &str,
    current_level: u32,
    levels: &BTreeMap<u32, RawListLevel>,
) -> Option<(NumberingPattern, bool)> {
    let mut pieces: Vec<(String, char)> = Vec::new();
    let mut literal: String = String::new();
    let mut chars = level_text.chars().peekable();
    let mut saw_current_level: bool = false;
    let mut saw_parent_level: bool = false;

    while let Some(ch) = chars.next() {
        if ch == '%' {
            let mut digits: String = String::new();
            while let Some(next) = chars.peek().copied() {
                if next.is_ascii_digit() {
                    digits.push(next);
                    chars.next();
                } else {
                    break;
                }
            }

            if digits.is_empty() {
                literal.push(ch);
                continue;
            }

            let referenced_level: u32 = digits.parse::<u32>().ok()?.checked_sub(1)?;
            let referenced = levels.get(&referenced_level)?;
            let symbol = if levels.get(&current_level)?.uses_legal_numbering {
                '1'
            } else {
                counter_symbol(&referenced.number_format)?
            };
            pieces.push((std::mem::take(&mut literal), symbol));
            if referenced_level == current_level {
                saw_current_level = true;
            } else if referenced_level < current_level {
                saw_parent_level = true;
            }
            continue;
        }

        literal.push(ch);
    }

    // A lvlText without placeholders is static text on every item, as
    // `format_inline_number` prints it; only a label of other levels gains
    // the current counter.
    if !saw_current_level && !pieces.is_empty() {
        let current = levels.get(&current_level)?;
        let symbol = if current.uses_legal_numbering {
            '1'
        } else {
            counter_symbol(&current.number_format)?
        };
        pieces.insert(0, (String::new(), symbol));
    }

    Some((
        NumberingPattern {
            pieces,
            suffix: literal,
        },
        saw_parent_level,
    ))
}

fn extract_raw_level(
    level: &docx_rs::Level,
    style_map: &super::StyleMap,
    theme_fonts: &super::ThemeFonts,
) -> RawListLevel {
    // Word allows restart thresholds 0..=7. Native probes above that range
    // retain child markers and count physical items only, as a zero does.
    let restart_trigger: Option<u32> =
        level
            .level_restart
            .as_ref()
            .and_then(serialize_u32)
            .map(|trigger| {
                if level.level > 0 && trigger > 7 {
                    tracing::debug!(
                        level = level.level,
                        restart_trigger = trigger,
                        normalized_restart_trigger = 0,
                        "Normalized out-of-range DOCX child restart threshold"
                    );
                    0
                } else {
                    trigger
                }
            });
    let mut paragraph_style = super::text::extract_paragraph_style(&level.paragraph_property);
    paragraph_style.numbering_suffix = Some(match level.suffix {
        docx_rs::LevelSuffixType::Tab => ListMarkerSuffix::Tab,
        docx_rs::LevelSuffixType::Space => ListMarkerSuffix::Space,
        docx_rs::LevelSuffixType::Nothing => ListMarkerSuffix::Nothing,
    });
    // Level run properties format the generated marker, not the body runs.
    // Resolve its character style/theme like ordinary text; unspecified
    // properties inherit surrounding list text at render time.
    let marker_style =
        super::resolve_run_style(&level.run_property, false, None, style_map, theme_fonts);
    RawListLevel {
        start: serialize_u32(&level.start).unwrap_or(1),
        number_format: level.format.val.clone(),
        level_text: serialize_string(&level.text).unwrap_or_default(),
        paragraph_style,
        marker_style: (marker_style != TextStyle::default()).then_some(marker_style),
        has_start_override: false,
        uses_legal_numbering: level.is_lgl.is_some(),
        restart_trigger,
    }
}

fn resolve_numbering(
    num: &docx_rs::Numbering,
    numberings: &docx_rs::Numberings,
    style_map: &super::StyleMap,
    theme_fonts: &super::ThemeFonts,
) -> ResolvedNumbering {
    let abstract_num = numberings
        .abstract_nums
        .iter()
        .find(|abstract_num| abstract_num.id == num.abstract_num_id);

    let mut raw_levels: BTreeMap<u32, RawListLevel> = abstract_num
        .map(|abstract_num| {
            abstract_num
                .levels
                .iter()
                .map(|level| {
                    (
                        level.level as u32,
                        extract_raw_level(level, style_map, theme_fonts),
                    )
                })
                .collect()
        })
        .unwrap_or_default();

    for override_level in &num.level_overrides {
        let level_index = override_level.level as u32;
        if let Some(level) = &override_level.override_level {
            let abstract_restart_trigger = raw_levels
                .get(&level_index)
                .and_then(|level| level.restart_trigger);
            let mut raw_level = extract_raw_level(level, style_map, theme_fonts);
            raw_level.has_start_override = true;
            // Word ignores lvlRestart in complete level overrides (MS-OI29500
            // 2.1.282); the abstract policy remains in force.
            raw_level.restart_trigger = abstract_restart_trigger;
            raw_levels.insert(level_index, raw_level);
        }
        if let Some(start) = override_level.override_start {
            raw_levels
                .entry(level_index)
                .and_modify(|level| {
                    level.start = start as u32;
                    level.has_start_override = true;
                })
                .or_insert_with(|| RawListLevel {
                    start: start as u32,
                    number_format: "decimal".to_string(),
                    level_text: format!("%{}.", level_index + 1),
                    paragraph_style: ParagraphStyle::default(),
                    marker_style: None,
                    has_start_override: true,
                    uses_legal_numbering: false,
                    restart_trigger: None,
                });
        }
    }

    let levels: BTreeMap<u32, ResolvedListLevel> = raw_levels
        .iter()
        .map(|(level_index, level)| {
            let kind = level_kind(&level.number_format);
            let (numbering_pattern, full_numbering) = if kind == ListKind::Ordered {
                build_numbering_pattern(&level.level_text, *level_index, &raw_levels)
                    .map(|(pattern, full)| (Some(pattern), full))
                    .unwrap_or((None, false))
            } else {
                (None, false)
            };

            (
                *level_index,
                ResolvedListLevel {
                    level_text: level.level_text.clone(),
                    number_format: level.number_format.clone(),
                    uses_legal_numbering: level.uses_legal_numbering,
                    style: ListLevelStyle {
                        kind,
                        numbering_pattern,
                        full_numbering,
                        // Word renders the numbering definition's lvlText
                        // glyph per level (•, ○, ▪ …); dropping it made
                        // every level reuse the level-1 disc (issue #356).
                        marker_text: (kind == ListKind::Unordered)
                            .then(|| level.level_text.clone())
                            .filter(|text| !text.is_empty()),
                        marker_style: level.marker_style.clone(),
                    },
                    paragraph_style: level.paragraph_style.clone(),
                    start: level.start,
                    has_start_override: level.has_start_override,
                    restart_trigger: level.restart_trigger,
                },
            )
        })
        .collect();

    ResolvedNumbering {
        abstract_num_id: num.abstract_num_id,
        levels,
    }
}

pub(super) fn build_numbering_map(
    numberings: &docx_rs::Numberings,
    style_map: &super::StyleMap,
    theme_fonts: &super::ThemeFonts,
) -> NumberingMap {
    numberings
        .numberings
        .iter()
        .map(|numbering| {
            (
                numbering.id,
                resolve_numbering(numbering, numberings, style_map, theme_fonts),
            )
        })
        .collect()
}

/// Extract numbering info from a paragraph, if it has numPr.
///
/// Falls back to the style definition when the paragraph carries no inline
/// `<w:numPr>` — this handles Word built-in list styles such as "List Bullet"
/// and "List Number" whose numbering is defined on the style, not the paragraph.
pub(super) fn extract_num_info(
    para: &docx_rs::Paragraph,
    styles: &docx_rs::Styles,
) -> Option<NumInfo> {
    if para.has_numbering {
        let numbering_property = para.property.numbering_property.as_ref()?;
        let num_id = numbering_property.id.as_ref()?.id;
        let level = numbering_property
            .level
            .as_ref()
            .map_or(0, |level| level.val as u32);
        if num_id == 0 {
            return None;
        }
        return Some(NumInfo {
            num_id,
            level,
            style_supplied_indents: StyleSuppliedIndents::default(),
        });
    }

    let style_id = para
        .property
        .style
        .as_ref()
        .map(|style| style.val.as_str())?;
    let style = styles
        .styles
        .iter()
        .find(|style| style.style_id == style_id)?;
    let numbering_property = style.paragraph_property.numbering_property.as_ref()?;
    let num_id = numbering_property.id.as_ref()?.id;
    let level = numbering_property
        .level
        .as_ref()
        .map_or(0, |level| level.val as u32);
    if num_id == 0 {
        return None;
    }
    Some(NumInfo {
        num_id,
        level,
        style_supplied_indents: StyleSuppliedIndents::default(),
    })
}

/// An intermediate element that carries optional numbering info alongside blocks.
pub(super) enum TaggedElement {
    /// A regular block (non-list paragraph, table, image, page break, etc.)
    Plain(Vec<Block>),
    /// A list paragraph with its numbering info and the paragraph IR.
    ///
    /// Boxed because a `Paragraph` is an order of magnitude larger than the
    /// `Plain` variant's vector, and every element of the stream would
    /// otherwise be sized for the larger one.
    ListParagraph {
        info: NumInfo,
        paragraph: Box<Paragraph>,
    },
}

/// A list item paired with the `numId` of the paragraph it came from, so a
/// merged list can resolve per-item numbering across differing `numId`s.
struct NumberedItem {
    num_id: usize,
    series: NumberingSeries,
    item: ListItem,
}

/// Format each Word placeholder from its referenced counter and declared format.
/// A legal label decimalizes all its references, independently of their own flags.
fn format_inline_number(
    current_level: u32,
    numbering: &ResolvedNumbering,
    counters: &BTreeMap<u32, NumberingCounter>,
) -> String {
    let level = &numbering.levels[&current_level];
    let mut label: String = String::new();
    let mut characters = level.level_text.chars().peekable();
    while let Some(character) = characters.next() {
        if character != '%' {
            label.push(character);
            continue;
        }
        let mut digits: String = String::new();
        while let Some(digit) = characters.peek().copied().filter(char::is_ascii_digit) {
            digits.push(digit);
            characters.next();
        }
        let referenced_level = digits
            .parse::<u32>()
            .ok()
            .and_then(|index| index.checked_sub(1));
        let referenced = referenced_level
            .and_then(|index| numbering.levels.get(&index).map(|level| (index, level)));
        if let Some((index, referenced)) = referenced {
            let value: u32 = counters.get(&index).map_or(0, |counter| counter.value);
            let number_format: &str = if level.uses_legal_numbering {
                "decimal"
            } else {
                &referenced.number_format
            };
            label.push_str(&format_inline_counter(value, number_format));
        } else {
            // An invalid or undefined reference does not authorize replacing
            // another level's counter or swallowing literal source text.
            label.push('%');
            label.push_str(&digits);
        }
    }
    label
}

fn format_inline_counter(value: u32, number_format: &str) -> String {
    match number_format {
        "decimalZero" => format!("{value:02}"),
        "lowerLetter" => crate::numbering::format_word_alphabetic_counter(value, false),
        "upperLetter" => crate::numbering::format_word_alphabetic_counter(value, true),
        "lowerRoman" => crate::numbering::format_roman_counter(value, false),
        "upperRoman" => crate::numbering::format_roman_counter(value, true),
        _ => value.to_string(),
    }
}

fn numbering_series(num_id: usize, numberings: &NumberingMap) -> NumberingSeries {
    numberings
        .get(&num_id)
        .map(|numbering| NumberingSeries::Abstract(numbering.abstract_num_id))
        .unwrap_or(NumberingSeries::Numbering(num_id))
}

/// A one-based trigger resets the child when its named ancestor or an earlier
/// ancestor is used; zero never resets.
fn reset_deeper_numbering_counters(
    counters: &mut BTreeMap<u32, NumberingCounter>,
    used_level: u32,
    numbering: Option<&ResolvedNumbering>,
) {
    counters.retain(|counter_level, _| {
        if *counter_level <= used_level {
            return true;
        }
        let restart_trigger: Option<u32> = numbering
            .and_then(|numbering| numbering.levels.get(counter_level))
            .and_then(|level| level.restart_trigger);
        let should_reset: bool = match restart_trigger {
            Some(0) => false,
            Some(trigger) => used_level < trigger,
            None => true,
        };
        tracing::trace!(
            used_level,
            counter_level = *counter_level,
            restart_trigger = ?restart_trigger,
            should_reset,
            "Resolved DOCX child-level numbering restart"
        );
        !should_reset
    });
}

// Word counts a missing resettable ancestor before its descendant. Never-
// restarting ancestors initialize only on physical items, as native probes show.
fn initialize_implicit_numbering_ancestors(
    counters: &mut BTreeMap<u32, NumberingCounter>,
    used_level: u32,
    numbering: &ResolvedNumbering,
) {
    for (ancestor_level, level) in numbering.levels.range(..used_level) {
        if level.style.kind != ListKind::Ordered || level.restart_trigger == Some(0) {
            continue;
        }
        counters.entry(*ancestor_level).or_insert_with(|| {
            tracing::trace!(
                used_level,
                ancestor_level,
                start = level.start,
                "Initialized implicit DOCX numbering ancestor"
            );
            NumberingCounter {
                value: level.start,
                is_implicit: true,
            }
        });
    }
}

fn apply_numbering_level_indentation(
    paragraph: &mut Paragraph,
    level_style: Option<&ParagraphStyle>,
) {
    let Some(level_style) = level_style else {
        return;
    };

    let style = &mut paragraph.style;
    style.numbering_suffix = level_style.numbering_suffix;
    style.indent_left = style.indent_left.or(level_style.indent_left);
    style.indent_right = style.indent_right.or(level_style.indent_right);
    style.indent_first_line = style.indent_first_line.or(level_style.indent_first_line);
    if style.tab_stops.is_none() {
        style.tab_stops = level_style.tab_stops.clone();
    }
}

/// Whether a style-chain indent stands in for one the paragraph's direct
/// numbering level declares. Word uses the level's there; the parser keeps
/// the style's (#2383).
fn style_indent_masks_level(
    supplied: StyleSuppliedIndents,
    level_style: Option<&ParagraphStyle>,
) -> bool {
    level_style.is_some_and(|level| {
        (supplied.left && level.indent_left.is_some())
            || (supplied.first_line && level.indent_first_line.is_some())
    })
}

/// Clamps a list paragraph's left indent and marker origin at the margin, the
/// geometry the renderer drew for every list before list indents kept their
/// sign (issue #2233).
fn clamp_list_indent_at_margin(style: &mut ParagraphStyle) {
    let Some(left) = style.indent_left.map(|left| left.max(0.0)) else {
        return;
    };
    style.indent_left = Some(left);
    if let Some(first_line) = style
        .indent_first_line
        .filter(|first_line| *first_line < 0.0)
    {
        style.indent_first_line = Some((left + first_line).max(0.0) - left);
    }
}

fn finalize_list(numbered_items: Vec<NumberedItem>, numberings: &NumberingMap) -> List {
    // Build merged per-level styles from every numId present. The first item
    // encountered at a given level establishes that level's style — adjacent
    // list paragraphs authored with different numIds (common in pandoc/
    // LibreOffice output, issue #176) thus share one coherent style map.
    let mut level_styles: BTreeMap<u32, ListLevelStyle> = BTreeMap::new();
    for numbered in &numbered_items {
        if let Some(resolved) = numberings.get(&numbered.num_id)
            && let Some(resolved_level) = resolved.levels.get(&numbered.item.level)
        {
            level_styles
                .entry(numbered.item.level)
                .or_insert_with(|| resolved_level.style.clone());
        }
    }

    // The overall list kind follows level 0 (or the shallowest level present).
    let kind = level_styles
        .get(&0)
        .map(|style| style.kind)
        .or_else(|| level_styles.values().next().map(|style| style.kind))
        .unwrap_or(ListKind::Unordered);

    let items: Vec<ListItem> = numbered_items
        .into_iter()
        .map(|numbered| numbered.item)
        .collect();

    List {
        kind,
        items,
        level_styles,
    }
}

/// Group consecutive list paragraphs into List blocks. Compatible adjacent list
/// paragraphs are merged even when their `numId` differs, so ordered numbering
/// continues and `ilvl` nesting is preserved (issue #176). A top-level change
/// between ordered and unordered numbering starts a new list, as does any
/// non-list element.
pub(super) fn group_into_lists(
    elements: Vec<TaggedElement>,
    numberings: &NumberingMap,
) -> Vec<Block> {
    let mut result: Vec<Block> = Vec::new();
    let mut current_list: Vec<NumberedItem> = Vec::new();
    let mut counters: HashMap<NumberingSeries, BTreeMap<u32, NumberingCounter>> = HashMap::new();
    let mut last_num_id: HashMap<NumberingSeries, usize> = HashMap::new();

    for element in elements {
        match element {
            TaggedElement::ListParagraph { info, paragraph } => {
                let series = numbering_series(info.num_id, numberings);
                let resolved_numbering = numberings.get(&info.num_id);
                let resolved_level =
                    resolved_numbering.and_then(|numbering| numbering.levels.get(&info.level));
                let is_ordered =
                    resolved_level.is_some_and(|level| level.style.kind == ListKind::Ordered);
                let is_num_id_change = last_num_id
                    .get(&series)
                    .is_some_and(|previous| *previous != info.num_id);
                let has_explicit_restart = resolved_level
                    .is_some_and(|level| level.has_start_override && is_num_id_change);
                let changes_series = current_list
                    .last()
                    .is_some_and(|numbered| numbered.series != series);
                let current_root = current_list
                    .iter()
                    .filter_map(|numbered| {
                        numberings
                            .get(&numbered.num_id)
                            .and_then(|numbering| numbering.levels.get(&numbered.item.level))
                            .map(|level| (numbered.item.level, level.style.kind))
                    })
                    .min_by_key(|(level, _)| *level);
                let starts_new_list = changes_series
                    && current_root.is_some_and(|(root_level, root_kind)| {
                        info.level <= root_level
                            && resolved_level.is_some_and(|level| level.style.kind != root_kind)
                    });
                // A continuation can start at a child after a plain paragraph.
                // A later shallower item needs its own root style and indent.
                let returns_to_shallower_root: bool = current_list
                    .first()
                    .is_some_and(|numbered| info.level < numbered.item.level);
                if starts_new_list || returns_to_shallower_root {
                    result.push(Block::List(finalize_list(
                        std::mem::take(&mut current_list),
                        numberings,
                    )));
                }
                let is_first_in_block = current_list.is_empty();
                let mut paragraph = paragraph;
                if let Some(series_counters) = counters.get_mut(&series) {
                    reset_deeper_numbering_counters(
                        series_counters,
                        info.level,
                        resolved_numbering,
                    );
                }

                if is_ordered {
                    initialize_implicit_numbering_ancestors(
                        counters.entry(series).or_default(),
                        info.level,
                        resolved_numbering.expect("ordered numbering must be resolved"),
                    );
                }

                // Within Word's documented 0..=7 range, same/later triggers
                // suppress a child level's marker and numbering indent. Root
                // flags are ignored; paragraph-owned formatting remains intact.
                let suppresses_marker: bool = info.level > 0
                    && resolved_level.is_some_and(|level| {
                        level.restart_trigger.is_some_and(|trigger| {
                            (1..=7).contains(&trigger) && trigger >= info.level
                        })
                    });
                if suppresses_marker {
                    if is_ordered {
                        let level = resolved_level.expect("ordered level must be resolved");
                        let series_counters = counters.entry(series).or_default();
                        let number: u32 = if has_explicit_restart {
                            level.start
                        } else {
                            series_counters
                                .get(&info.level)
                                .map_or(level.start, |counter| counter.value.saturating_add(1))
                        };
                        series_counters.insert(
                            info.level,
                            NumberingCounter {
                                value: number,
                                is_implicit: false,
                            },
                        );
                    }
                    last_num_id.insert(series, info.num_id);
                    if !current_list.is_empty() {
                        result.push(Block::List(finalize_list(
                            std::mem::take(&mut current_list),
                            numberings,
                        )));
                    }
                    tracing::trace!(
                        num_id = info.num_id,
                        level = info.level,
                        restart_trigger = ?resolved_level.and_then(|level| level.restart_trigger),
                        "Suppressed DOCX numbering marker and level indentation"
                    );
                    result.push(Block::Paragraph(*paragraph));
                    continue;
                }

                // Word puts numbers with `w:ind left=0 hanging=0` inline:
                // number, source separator, first-line text, and continuation
                // lines wrap back to the margin. A hanging-indent list would
                // indent every continuation line instead (issue #357).
                let is_flush_numbered = is_ordered
                    && resolved_level.is_some_and(|level| {
                        level.paragraph_style.indent_left == Some(0.0)
                            && level.paragraph_style.indent_first_line == Some(0.0)
                    })
                    && paragraph.style.indent_left.unwrap_or(0.0) == 0.0
                    && paragraph.style.indent_first_line.unwrap_or(0.0) == 0.0;
                let has_paragraph_frame: bool =
                    paragraph.style.background.is_some() || paragraph.style.border.is_some();
                let is_inline_numbered: bool =
                    is_flush_numbered || (is_ordered && has_paragraph_frame);
                if is_inline_numbered {
                    let level = resolved_level.expect("ordered level must be resolved");
                    paragraph.style.has_inline_list_number = true;
                    paragraph.style.numbering_suffix = level.paragraph_style.numbering_suffix;
                    if !is_flush_numbered {
                        // The list renderer emits item runs without their paragraph frame.
                        // Inline numbering lets normal paragraph codegen keep shading and borders.
                        apply_numbering_level_indentation(
                            &mut paragraph,
                            resolved_level.map(|level| &level.paragraph_style),
                        );
                    }
                    let number_tab_position: f64 = paragraph.style.indent_left.unwrap_or(0.0);
                    if number_tab_position > 0.0 {
                        let mut tab_stops: Vec<TabStop> =
                            paragraph.style.tab_stops.take().unwrap_or_default();
                        tab_stops.insert(
                            0,
                            TabStop {
                                position: number_tab_position,
                                alignment: TabAlignment::Left,
                                leader: TabLeader::None,
                            },
                        );
                        paragraph.style.tab_stops = Some(tab_stops);
                    }
                    let series_counters = counters.entry(series).or_default();
                    let should_restart =
                        has_explicit_restart || !series_counters.contains_key(&info.level);
                    let number = if should_restart {
                        level.start
                    } else {
                        series_counters[&info.level].value.saturating_add(1)
                    };
                    series_counters.insert(
                        info.level,
                        NumberingCounter {
                            value: number,
                            is_implicit: false,
                        },
                    );
                    last_num_id.insert(series, info.num_id);

                    if !current_list.is_empty() {
                        result.push(Block::List(finalize_list(
                            std::mem::take(&mut current_list),
                            numberings,
                        )));
                    }
                    let label = format_inline_number(
                        info.level,
                        resolved_numbering.expect("ordered numbering must be resolved"),
                        series_counters,
                    );
                    // The paragraph mark formats the number, not the first
                    // run (issue #2337).
                    let mut prefix_style = paragraph
                        .style
                        .numbering_label_style
                        .as_deref()
                        .or_else(|| paragraph.runs.first().map(|run| &run.style))
                        .cloned()
                        .unwrap_or_default();
                    if let Some(marker_style) = &level.style.marker_style {
                        prefix_style.merge_from(marker_style);
                    }
                    let suffix = level
                        .paragraph_style
                        .numbering_suffix
                        .unwrap_or(ListMarkerSuffix::Tab);
                    if suffix == ListMarkerSuffix::Space {
                        // Word's generated separator uses Arial at the
                        // marker's size, independently of its label family.
                        paragraph.runs.insert(
                            0,
                            Run {
                                text: " ".to_string(),
                                style: crate::ir::TextStyle {
                                    font_family: Some("Arial".to_string()),
                                    font_size: prefix_style.font_size,
                                    ..crate::ir::TextStyle::default()
                                },
                                href: None,
                                footnote: None,
                                inline_box: None,
                            },
                        );
                    }
                    paragraph.runs.insert(
                        0,
                        Run {
                            text: if suffix == ListMarkerSuffix::Tab {
                                format!("{label}\t")
                            } else {
                                label
                            },
                            style: prefix_style,
                            href: None,
                            footnote: None,
                            inline_box: None,
                        },
                    );
                    tracing::trace!(
                        num_id = info.num_id,
                        level = info.level,
                        ?suffix,
                        "Rendered DOCX inline numbering separator"
                    );
                    result.push(Block::Paragraph(*paragraph));
                    continue;
                }

                apply_numbering_level_indentation(
                    &mut paragraph,
                    resolved_level.map(|level| &level.paragraph_style),
                );
                // TODO(#2383): Word lays such a paragraph out at its level's
                // indent. Until the parser does, keep it at the margin as
                // before: hanging the style-chain indent into the margin would
                // move it further from Word (issue #2233).
                if style_indent_masks_level(
                    info.style_supplied_indents,
                    resolved_level.map(|level| &level.paragraph_style),
                ) {
                    clamp_list_indent_at_margin(&mut paragraph.style);
                }
                let mut item = ListItem {
                    content: vec![*paragraph],
                    level: info.level,
                    start_at: None,
                };
                if is_ordered {
                    let level = resolved_level.expect("ordered level must be resolved");
                    let series_counters = counters.entry(series).or_default();
                    let should_restart =
                        has_explicit_restart || !series_counters.contains_key(&info.level);
                    let number = if should_restart {
                        level.start
                    } else {
                        series_counters[&info.level].value.saturating_add(1)
                    };
                    let was_implicit: bool = series_counters
                        .get(&info.level)
                        .is_some_and(|counter| counter.is_implicit);
                    series_counters.insert(
                        info.level,
                        NumberingCounter {
                            value: number,
                            is_implicit: false,
                        },
                    );
                    if should_restart
                        || is_first_in_block
                        || changes_series
                        || level.restart_trigger.is_some()
                        || was_implicit
                    {
                        item.start_at = Some(number);
                    }
                }
                last_num_id.insert(series, info.num_id);
                current_list.push(NumberedItem {
                    num_id: info.num_id,
                    series,
                    item,
                });
            }
            TaggedElement::Plain(blocks) => {
                if !current_list.is_empty() {
                    result.push(Block::List(finalize_list(
                        std::mem::take(&mut current_list),
                        numberings,
                    )));
                }
                result.extend(blocks);
            }
        }
    }

    if !current_list.is_empty() {
        result.push(Block::List(finalize_list(current_list, numberings)));
    }

    result
}
