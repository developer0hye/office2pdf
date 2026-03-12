use std::collections::{BTreeMap, HashMap};

use crate::ir::{Block, List, ListItem, ListKind, ListLevelStyle, Paragraph, TabStop};

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
    indent_left: Option<f64>,
    indent_right: Option<f64>,
    indent_first_line: Option<f64>,
    tab_stops: Option<Vec<TabStop>>,
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
    indent_left: Option<f64>,
    indent_right: Option<f64>,
    indent_first_line: Option<f64>,
    tab_stops: Option<Vec<TabStop>>,
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
    let indent = level.paragraph_property.indent.as_ref();
    let indent_left = indent
        .and_then(|value| value.start)
        .map(|value| value as f64 / 20.0);
    let indent_right = indent
        .and_then(|value| value.end)
        .map(|value| value as f64 / 20.0);
    let indent_first_line = indent.and_then(|value| {
        value
            .special_indent
            .map(|special_indent| match special_indent {
                docx_rs::SpecialIndentType::FirstLine(amount) => amount as f64 / 20.0,
                docx_rs::SpecialIndentType::Hanging(amount) => -(amount as f64 / 20.0),
            })
    });

    RawListLevel {
        start: serialize_u32(&level.start).unwrap_or(1),
        number_format: level.format.val.clone(),
        level_text: serialize_string(&level.text).unwrap_or_default(),
        indent_left,
        indent_right,
        indent_first_line,
        tab_stops: super::text::extract_tab_stops(&level.paragraph_property.tabs),
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
                    indent_left: None,
                    indent_right: None,
                    indent_first_line: None,
                    tab_stops: None,
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
                    indent_left: level.indent_left,
                    indent_right: level.indent_right,
                    indent_first_line: level.indent_first_line,
                    tab_stops: level.tab_stops.clone(),
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

#[derive(Debug, Clone)]
struct PendingListItem {
    num_id: usize,
    level: u32,
    paragraph: Paragraph,
}

#[derive(Debug, Clone)]
struct PendingList {
    root_num_id: usize,
    root_level: u32,
    items: Vec<PendingListItem>,
}

impl PendingList {
    fn new(info: NumInfo, paragraph: Paragraph) -> Self {
        Self {
            root_num_id: info.num_id,
            root_level: info.level,
            items: vec![PendingListItem {
                num_id: info.num_id,
                level: info.level,
                paragraph,
            }],
        }
    }

    fn push(&mut self, info: NumInfo, paragraph: Paragraph) {
        self.items.push(PendingListItem {
            num_id: info.num_id,
            level: info.level,
            paragraph,
        });
    }
}

fn resolved_list_level(
    numberings: &NumberingMap,
    num_id: usize,
    level: u32,
) -> Option<&ResolvedListLevel> {
    numberings.get(&num_id)?.levels.get(&level)
}

fn resolved_level_kind(numberings: &NumberingMap, num_id: usize, level: u32) -> Option<ListKind> {
    resolved_list_level(numberings, num_id, level)
        .map(|resolved| resolved.style.kind)
        .or_else(|| numberings.get(&num_id).map(|resolved| resolved.kind))
}

fn fallback_level_style(kind: ListKind) -> ListLevelStyle {
    ListLevelStyle {
        kind,
        numbering_pattern: None,
        full_numbering: false,
        marker_text: None,
        marker_style: None,
    }
}

fn list_belongs_to_pending(current: &PendingList, info: &NumInfo) -> bool {
    if info.level < current.root_level {
        return false;
    }

    if info.level == current.root_level {
        return info.num_id == current.root_num_id;
    }

    current.items.iter().any(|item| item.level < info.level)
}

fn finalize_list(pending: PendingList, numberings: &NumberingMap) -> List {
    let root_kind: ListKind = pending
        .items
        .iter()
        .find(|item| item.level == pending.root_level)
        .and_then(|item| resolved_level_kind(numberings, item.num_id, item.level))
        .unwrap_or(ListKind::Unordered);

    let mut level_styles: BTreeMap<u32, ListLevelStyle> = BTreeMap::new();
    for item in &pending.items {
        level_styles.entry(item.level).or_insert_with(|| {
            resolved_list_level(numberings, item.num_id, item.level)
                .map(|resolved| resolved.style.clone())
                .or_else(|| {
                    resolved_level_kind(numberings, item.num_id, item.level)
                        .map(fallback_level_style)
                })
                .unwrap_or_else(|| fallback_level_style(root_kind))
        });
    }

    let mut items: Vec<ListItem> = Vec::with_capacity(pending.items.len());
    let mut previous_level: Option<u32> = None;
    let mut previous_num_id: Option<usize> = None;
    for pending_item in pending.items {
        let PendingListItem {
            num_id,
            level,
            mut paragraph,
        } = pending_item;
        let style_kind: ListKind = level_styles
            .get(&level)
            .map(|style| style.kind)
            .unwrap_or(root_kind);
        let resolved_level = resolved_list_level(numberings, num_id, level);
        let start_value: u32 = resolved_level.map(|resolved| resolved.start).unwrap_or(1);

        if let Some(resolved_level) = resolved_level {
            if paragraph.style.indent_left.is_none() {
                paragraph.style.indent_left = resolved_level.indent_left;
            }
            if paragraph.style.indent_right.is_none() {
                paragraph.style.indent_right = resolved_level.indent_right;
            }
            if paragraph.style.indent_first_line.is_none() {
                paragraph.style.indent_first_line = resolved_level.indent_first_line;
            }
            if paragraph.style.tab_stops.is_none() {
                paragraph.style.tab_stops = resolved_level.tab_stops.clone();
            }
        }

        let start_at: Option<u32> = if style_kind == ListKind::Ordered {
            match (previous_level, previous_num_id) {
                (None, _) => Some(start_value),
                (Some(prev_level), Some(prev_num_id))
                    if level > prev_level || (level == prev_level && num_id != prev_num_id) =>
                {
                    Some(start_value)
                }
                _ => None,
            }
        } else {
            None
        };

        items.push(ListItem {
            content: vec![paragraph],
            level,
            start_at,
        });
        previous_level = Some(level);
        previous_num_id = Some(num_id);
    }

    List {
        kind: root_kind,
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
    let mut current_list: Option<PendingList> = None;

    for element in elements {
        match element {
            TaggedElement::ListParagraph { info, paragraph } => {
                if let Some(current) = current_list.as_mut() {
                    if list_belongs_to_pending(current, &info) {
                        current.push(info, paragraph);
                        continue;
                    }
                    let finished = current_list
                        .take()
                        .expect("current list should exist when flushing");
                    result.push(Block::List(finalize_list(finished, numberings)));
                }
                current_list = Some(PendingList::new(info, paragraph));
            }
            TaggedElement::Plain(blocks) => {
                if let Some(list) = current_list.take() {
                    result.push(Block::List(finalize_list(list, numberings)));
                }
                result.extend(blocks);
            }
        }
    }

    if let Some(list) = current_list {
        result.push(Block::List(finalize_list(list, numberings)));
    }

    result
}
