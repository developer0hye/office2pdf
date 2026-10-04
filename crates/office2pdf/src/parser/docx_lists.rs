use std::collections::{BTreeMap, HashMap};

use crate::ir::{
    Block, List, ListItem, ListKind, ListLevelStyle, Paragraph, ParagraphStyle, Run, TabAlignment,
    TabLeader, TabStop,
};

/// Numbering info extracted from a paragraph's numPr.
#[derive(Debug, Clone)]
pub(super) struct NumInfo {
    pub(super) num_id: usize,
    pub(super) level: u32,
}

#[derive(Debug, Clone)]
struct ResolvedListLevel {
    style: ListLevelStyle,
    paragraph_style: ParagraphStyle,
    start: u32,
    has_start_override: bool,
    /// Raw Word `lvlText` (e.g. "제%1조") for inline-numbered paragraphs.
    level_text: String,
    number_format: String,
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
    has_start_override: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum NumberingSeries {
    Abstract(usize),
    Numbering(usize),
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

fn typst_counter_symbol(number_format: &str) -> Option<&'static str> {
    match number_format {
        "decimal" | "decimalZero" => Some("1"),
        "lowerLetter" => Some("a"),
        "upperLetter" => Some("A"),
        "lowerRoman" => Some("i"),
        "upperRoman" => Some("I"),
        _ => None,
    }
}

fn build_typst_numbering_pattern(
    level_text: &str,
    current_level: u32,
    levels: &BTreeMap<u32, RawListLevel>,
) -> Option<(String, bool)> {
    let mut pattern: String = String::new();
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
                pattern.push(ch);
                continue;
            }

            let referenced_level: u32 = digits.parse::<u32>().ok()?.checked_sub(1)?;
            let referenced = levels.get(&referenced_level)?;
            let symbol = typst_counter_symbol(&referenced.number_format)?;
            pattern.push_str(symbol);
            if referenced_level == current_level {
                saw_current_level = true;
            } else if referenced_level < current_level {
                saw_parent_level = true;
            }
            continue;
        }

        pattern.push(ch);
    }

    if !saw_current_level {
        let current = levels.get(&current_level)?;
        let symbol = typst_counter_symbol(&current.number_format)?;
        pattern.insert_str(0, symbol);
    }

    Some((pattern, saw_parent_level))
}

fn extract_raw_level(level: &docx_rs::Level) -> RawListLevel {
    RawListLevel {
        start: serialize_u32(&level.start).unwrap_or(1),
        number_format: level.format.val.clone(),
        level_text: serialize_string(&level.text).unwrap_or_default(),
        paragraph_style: super::text::extract_paragraph_style(&level.paragraph_property),
        has_start_override: false,
    }
}

fn resolve_numbering(
    num: &docx_rs::Numbering,
    numberings: &docx_rs::Numberings,
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
                .map(|level| (level.level as u32, extract_raw_level(level)))
                .collect()
        })
        .unwrap_or_default();

    for override_level in &num.level_overrides {
        let level_index = override_level.level as u32;
        if let Some(level) = &override_level.override_level {
            let mut raw_level = extract_raw_level(level);
            raw_level.has_start_override = true;
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
                    has_start_override: true,
                });
        }
    }

    let levels: BTreeMap<u32, ResolvedListLevel> = raw_levels
        .iter()
        .map(|(level_index, level)| {
            let kind = level_kind(&level.number_format);
            let (numbering_pattern, full_numbering) = if kind == ListKind::Ordered {
                build_typst_numbering_pattern(&level.level_text, *level_index, &raw_levels)
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
                        marker_style: None,
                    },
                    paragraph_style: level.paragraph_style.clone(),
                    start: level.start,
                    has_start_override: level.has_start_override,
                },
            )
        })
        .collect();

    ResolvedNumbering {
        abstract_num_id: num.abstract_num_id,
        levels,
    }
}

pub(super) fn build_numbering_map(numberings: &docx_rs::Numberings) -> NumberingMap {
    numberings
        .numberings
        .iter()
        .map(|numbering| (numbering.id, resolve_numbering(numbering, numberings)))
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
        return Some(NumInfo { num_id, level });
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
    Some(NumInfo { num_id, level })
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

/// Substitute the current counter into a Word `lvlText` placeholder
/// ("제%1조" → "제3조"). Only the level's own placeholder participates for
/// inline-numbered paragraphs.
fn format_inline_number(level_text: &str, number_format: &str, number: u32) -> String {
    let formatted: String = match number_format {
        "decimalZero" => format!("{number:02}"),
        _ => number.to_string(),
    };
    let mut label = level_text.to_string();
    for placeholder_level in 1..=9 {
        label = label.replace(&format!("%{placeholder_level}"), &formatted);
    }
    label
}

fn numbering_series(num_id: usize, numberings: &NumberingMap) -> NumberingSeries {
    numberings
        .get(&num_id)
        .map(|numbering| NumberingSeries::Abstract(numbering.abstract_num_id))
        .unwrap_or(NumberingSeries::Numbering(num_id))
}

fn apply_numbering_level_indentation(
    paragraph: &mut Paragraph,
    level_style: Option<&ParagraphStyle>,
) {
    let Some(level_style) = level_style else {
        return;
    };

    let style = &mut paragraph.style;
    style.indent_left = style.indent_left.or(level_style.indent_left);
    style.indent_right = style.indent_right.or(level_style.indent_right);
    style.indent_first_line = style.indent_first_line.or(level_style.indent_first_line);
    if style.tab_stops.is_none() {
        style.tab_stops = level_style.tab_stops.clone();
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
    let mut counters: HashMap<NumberingSeries, BTreeMap<u32, u32>> = HashMap::new();
    let mut last_num_id: HashMap<NumberingSeries, usize> = HashMap::new();

    for element in elements {
        match element {
            TaggedElement::ListParagraph { info, paragraph } => {
                let series = numbering_series(info.num_id, numberings);
                let resolved_level = numberings
                    .get(&info.num_id)
                    .and_then(|numbering| numbering.levels.get(&info.level));
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
                if starts_new_list {
                    result.push(Block::List(finalize_list(
                        std::mem::take(&mut current_list),
                        numberings,
                    )));
                }
                let is_first_in_block = current_list.is_empty();
                let mut paragraph = paragraph;

                // Word puts numbers with `w:ind left=0 hanging=0` inline:
                // number, suffix tab, first-line text, and continuation
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
                        series_counters[&info.level].saturating_add(1)
                    };
                    series_counters.insert(info.level, number);
                    series_counters.retain(|item_level, _| *item_level <= info.level);
                    last_num_id.insert(series, info.num_id);

                    if !current_list.is_empty() {
                        result.push(Block::List(finalize_list(
                            std::mem::take(&mut current_list),
                            numberings,
                        )));
                    }
                    let label =
                        format_inline_number(&level.level_text, &level.number_format, number);
                    let prefix_style = paragraph
                        .runs
                        .first()
                        .map(|run| run.style.clone())
                        .unwrap_or_default();
                    paragraph.runs.insert(
                        0,
                        Run {
                            text: format!("{label}\t"),
                            style: prefix_style,
                            href: None,
                            footnote: None,
                            inline_box: None,
                        },
                    );
                    result.push(Block::Paragraph(*paragraph));
                    continue;
                }

                apply_numbering_level_indentation(
                    &mut paragraph,
                    resolved_level.map(|level| &level.paragraph_style),
                );
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
                        series_counters[&info.level].saturating_add(1)
                    };
                    series_counters.insert(info.level, number);
                    series_counters.retain(|level, _| *level <= info.level);
                    if should_restart || is_first_in_block || changes_series {
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
