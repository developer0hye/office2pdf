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

    ResolvedNumbering { levels }
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
    // Try paragraph-level numPr first.
    if para.has_numbering {
        if let Some(np) = para.property.numbering_property.as_ref() {
            let num_id = np.id.as_ref()?.id;
            let level = np.level.as_ref().map_or(0, |l| l.val as u32);
            if num_id != 0 {
                return Some(NumInfo { num_id, level });
            }
        }
    }

    // Fall back to the style definition (e.g. "List Bullet", "List Number").
    // Word stores numPr on the style when no inline numPr is present.
    let style_id = para.property.style.as_ref().map(|s| s.val.as_str())?;
    let style = styles.styles.iter().find(|s| s.style_id == style_id)?;
    let np = style.paragraph_property.numbering_property.as_ref()?;
    let num_id = np.id.as_ref()?.id;
    let level = np.level.as_ref().map_or(0, |l| l.val as u32);
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

/// A list item paired with the `numId` of the paragraph it came from, so a
/// merged list can resolve per-item numbering across differing `numId`s.
struct NumberedItem {
    num_id: usize,
    item: ListItem,
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

    // Ordered levels restart at their configured `start` only when first seen or
    // re-entered at a deeper level; otherwise the counter continues across the
    // merged items so "1." then "2." is preserved instead of "1." then "1.".
    let mut items: Vec<ListItem> = Vec::with_capacity(numbered_items.len());
    let mut previous_level: Option<u32> = None;
    for NumberedItem { num_id, mut item } in numbered_items {
        let resolved_level = numberings
            .get(&num_id)
            .and_then(|numbering| numbering.levels.get(&item.level));
        item.start_at = match (resolved_level, previous_level) {
            (Some(level), None) if level.style.kind == ListKind::Ordered => Some(level.start),
            (Some(level), Some(previous_level))
                if level.style.kind == ListKind::Ordered && item.level > previous_level =>
            {
                Some(level.start)
            }
            _ => None,
        };
        previous_level = Some(item.level);
        items.push(item);
    }

    List {
        kind,
        items,
        level_styles,
    }
}

/// Group consecutive list paragraphs into List blocks. Adjacent list paragraphs
/// are merged into a single list even when their `numId` differs, so ordered
/// numbering continues and `ilvl` nesting is preserved (issue #176). Any
/// non-list element ends the current list.
pub(super) fn group_into_lists(
    elements: Vec<TaggedElement>,
    numberings: &NumberingMap,
) -> Vec<Block> {
    let mut result: Vec<Block> = Vec::new();
    let mut current_list: Vec<NumberedItem> = Vec::new();

    for element in elements {
        match element {
            TaggedElement::ListParagraph { info, paragraph } => {
                current_list.push(NumberedItem {
                    num_id: info.num_id,
                    item: ListItem {
                        content: vec![paragraph],
                        level: info.level,
                        start_at: None,
                    },
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
