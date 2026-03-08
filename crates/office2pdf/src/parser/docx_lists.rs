use std::collections::{BTreeMap, HashMap};

use crate::ir::{Block, List, ListItem, ListKind, ListLevelStyle, Paragraph};

/// Numbering info extracted from a paragraph's numPr.
#[derive(Debug, Clone)]
pub(super) struct NumInfo {
    pub(super) num_id: usize,
    pub(super) level: u32,
}

#[derive(Debug, Clone)]
struct ResolvedListLevel {
    style: ListLevelStyle,
    start: u32,
}

#[derive(Debug, Clone)]
pub(super) struct ResolvedNumbering {
    kind: ListKind,
    levels: BTreeMap<u32, ResolvedListLevel>,
}

#[derive(Debug, Clone)]
struct RawListLevel {
    start: u32,
    number_format: String,
    level_text: String,
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
            raw_levels.insert(level_index, extract_raw_level(level));
        }
        if let Some(start) = override_level.override_start {
            raw_levels
                .entry(level_index)
                .and_modify(|level| level.start = start as u32)
                .or_insert_with(|| RawListLevel {
                    start: start as u32,
                    number_format: "decimal".to_string(),
                    level_text: format!("%{}.", level_index + 1),
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
                    style: ListLevelStyle {
                        kind,
                        numbering_pattern,
                        full_numbering,
                        marker_text: None,
                        marker_style: None,
                    },
                    start: level.start,
                },
            )
        })
        .collect();

    let kind = levels
        .get(&0)
        .map(|level| level.style.kind)
        .or_else(|| levels.values().next().map(|level| level.style.kind))
        .unwrap_or(ListKind::Unordered);

    ResolvedNumbering { kind, levels }
}

pub(super) fn build_numbering_map(numberings: &docx_rs::Numberings) -> NumberingMap {
    numberings
        .numberings
        .iter()
        .map(|numbering| (numbering.id, resolve_numbering(numbering, numberings)))
        .collect()
}

/// Extract numbering info from a paragraph, if it has numPr.
pub(super) fn extract_num_info(para: &docx_rs::Paragraph) -> Option<NumInfo> {
    if !para.has_numbering {
        return None;
    }
    let numbering_property = para.property.numbering_property.as_ref()?;
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
    ListParagraph { info: NumInfo, paragraph: Paragraph },
}

fn finalize_list(num_id: usize, mut items: Vec<ListItem>, numberings: &NumberingMap) -> List {
    let resolved = numberings.get(&num_id);

    let kind = resolved
        .map(|numbering| numbering.kind)
        .unwrap_or(ListKind::Unordered);
    let level_styles = resolved
        .map(|numbering| {
            numbering
                .levels
                .iter()
                .map(|(level, resolved_level)| (*level, resolved_level.style.clone()))
                .collect()
        })
        .unwrap_or_default();

    let mut previous_level: Option<u32> = None;
    for item in &mut items {
        let level_style = resolved.and_then(|numbering| numbering.levels.get(&item.level));
        item.start_at = match (level_style, previous_level) {
            (Some(level), None) if level.style.kind == ListKind::Ordered => Some(level.start),
            (Some(level), Some(previous_level))
                if level.style.kind == ListKind::Ordered && item.level > previous_level =>
            {
                Some(level.start)
            }
            _ => None,
        };
        previous_level = Some(item.level);
    }

    List {
        kind,
        items,
        level_styles,
    }
}

/// Group consecutive list paragraphs (with the same numId) into List blocks.
/// Non-list elements pass through unchanged.
pub(super) fn group_into_lists(
    elements: Vec<TaggedElement>,
    numberings: &NumberingMap,
) -> Vec<Block> {
    let mut result: Vec<Block> = Vec::new();
    let mut current_list: Option<(usize, Vec<ListItem>)> = None;

    for element in elements {
        match element {
            TaggedElement::ListParagraph { info, paragraph } => {
                if let Some((current_num_id, ref mut items)) = current_list {
                    if info.num_id == current_num_id {
                        items.push(ListItem {
                            content: vec![paragraph],
                            level: info.level,
                            start_at: None,
                        });
                        continue;
                    }
                    result.push(Block::List(finalize_list(
                        current_num_id,
                        std::mem::take(items),
                        numberings,
                    )));
                }
                current_list = Some((
                    info.num_id,
                    vec![ListItem {
                        content: vec![paragraph],
                        level: info.level,
                        start_at: None,
                    }],
                ));
            }
            TaggedElement::Plain(blocks) => {
                if let Some((num_id, items)) = current_list.take() {
                    result.push(Block::List(finalize_list(num_id, items, numberings)));
                }
                result.extend(blocks);
            }
        }
    }

    if let Some((num_id, items)) = current_list {
        result.push(Block::List(finalize_list(num_id, items, numberings)));
    }

    result
}
