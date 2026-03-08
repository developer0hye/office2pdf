use std::cell::Cell;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{Read, Seek};

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};

/// Maximum nesting depth for tables-within-tables.  Deeper nesting is silently
/// truncated to prevent stack overflow on pathological documents.
const MAX_TABLE_DEPTH: usize = 64;
use crate::ir::{
    Alignment, Block, BorderLineStyle, BorderSide, CellBorder, CellVerticalAlign, Chart, Color,
    ColumnLayout, Document, FloatingImage, FloatingTextBox, FlowPage, HFInline, HeaderFooter,
    HeaderFooterParagraph, ImageData, ImageFormat, Insets, LineSpacing, List, ListItem, ListKind,
    ListLevelStyle, Margins, MathEquation, Page, PageSize, Paragraph, ParagraphStyle, Run,
    StyleSheet, TabAlignment, TabLeader, TabStop, Table, TableCell, TableRow, TextDirection,
    TextStyle, VerticalTextAlign, WrapMode,
};
use crate::parser::Parser;

/// Parser for DOCX (Office Open XML Word) documents.
pub struct DocxParser;

/// Map from relationship ID → PNG image bytes.
type ImageMap = HashMap<String, Vec<u8>>;

/// Map from relationship ID → hyperlink URL.
type HyperlinkMap = HashMap<String, String>;

/// Parsed header/footer assets addressed by relationship ID.
#[derive(Default)]
struct HeaderFooterAssets {
    headers: HashMap<String, HeaderFooter>,
    footers: HashMap<String, HeaderFooter>,
}

/// Build a lookup map from the DOCX's hyperlinks (reader-populated field).
/// The reader stores hyperlinks as `(rid, url, type)` in `docx.hyperlinks`.
fn build_hyperlink_map(docx: &docx_rs::Docx) -> HyperlinkMap {
    docx.hyperlinks
        .iter()
        .map(|(rid, url, _type)| (rid.clone(), url.clone()))
        .collect()
}

fn scan_header_footer_relationships(
    rels_xml: &str,
) -> (HashMap<String, String>, HashMap<String, String>) {
    let mut headers: HashMap<String, String> = HashMap::new();
    let mut footers: HashMap<String, String> = HashMap::new();
    let mut reader = quick_xml::Reader::from_str(rels_xml);

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                if e.local_name().as_ref() != b"Relationship" {
                    continue;
                }

                let mut id: Option<String> = None;
                let mut target: Option<String> = None;
                let mut rel_type: Option<String> = None;

                for attr in e.attributes().flatten() {
                    match attr.key.local_name().as_ref() {
                        b"Id" => {
                            if let Ok(value) = attr.unescape_value() {
                                id = Some(value.to_string());
                            }
                        }
                        b"Target" => {
                            if let Ok(value) = attr.unescape_value() {
                                target = Some(value.to_string());
                            }
                        }
                        b"Type" => {
                            if let Ok(value) = attr.unescape_value() {
                                rel_type = Some(value.to_string());
                            }
                        }
                        _ => {}
                    }
                }

                let Some(id) = id else { continue };
                let Some(target) = target else { continue };
                let Some(rel_type) = rel_type else { continue };

                let full_path = if let Some(stripped) = target.strip_prefix('/') {
                    stripped.to_string()
                } else {
                    format!("word/{target}")
                };

                if rel_type.ends_with("/header") {
                    headers.insert(id, full_path);
                } else if rel_type.ends_with("/footer") {
                    footers.insert(id, full_path);
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    (headers, footers)
}

fn build_header_footer_assets<R: Read + Seek>(
    archive: &mut zip::ZipArchive<R>,
) -> HeaderFooterAssets {
    let rels_xml = match read_zip_text(archive, "word/_rels/document.xml.rels") {
        Some(xml) => xml,
        None => return HeaderFooterAssets::default(),
    };
    let (header_rels, footer_rels) = scan_header_footer_relationships(&rels_xml);
    let mut assets = HeaderFooterAssets::default();

    for (rid, path) in header_rels {
        let Some(xml) = read_zip_text(archive, &path) else {
            continue;
        };
        let Ok(header) = <docx_rs::Header as docx_rs::FromXML>::from_xml(xml.as_bytes()) else {
            continue;
        };
        if let Some(converted) = convert_docx_header(&header) {
            assets.headers.insert(rid, converted);
        }
    }

    for (rid, path) in footer_rels {
        let Some(xml) = read_zip_text(archive, &path) else {
            continue;
        };
        let Ok(footer) = <docx_rs::Footer as docx_rs::FromXML>::from_xml(xml.as_bytes()) else {
            continue;
        };
        if let Some(converted) = convert_docx_footer(&footer) {
            assets.footers.insert(rid, converted);
        }
    }

    assets
}

/// Build a lookup map from the DOCX's embedded images.
/// docx-rs converts all images to PNG; we use the PNG bytes.
fn build_image_map(docx: &docx_rs::Docx) -> ImageMap {
    docx.images
        .iter()
        .map(|(id, _path, _image, png)| (id.clone(), png.0.clone()))
        .collect()
}

/// Convert EMU (English Metric Units) to points.
/// 1 inch = 914400 EMU, 1 inch = 72 points, so 1 pt = 12700 EMU.
fn emu_to_pt(emu: u32) -> f64 {
    emu as f64 / 12700.0
}

/// Numbering info extracted from a paragraph's numPr.
#[derive(Debug, Clone)]
struct NumInfo {
    num_id: usize,
    level: u32,
}

#[derive(Debug, Clone)]
struct ResolvedListLevel {
    style: ListLevelStyle,
    start: u32,
}

#[derive(Debug, Clone)]
struct ResolvedNumbering {
    kind: ListKind,
    levels: BTreeMap<u32, ResolvedListLevel>,
}

#[derive(Debug, Clone)]
struct RawListLevel {
    start: u32,
    number_format: String,
    level_text: String,
}

type NumberingMap = HashMap<usize, ResolvedNumbering>;

fn serialize_string<T: serde::Serialize>(value: &T) -> Option<String> {
    serde_json::to_value(value)
        .ok()?
        .as_str()
        .map(|s| s.to_string())
}

fn serialize_u32<T: serde::Serialize>(value: &T) -> Option<u32> {
    serde_json::to_value(value)
        .ok()?
        .as_u64()
        .and_then(|v| u32::try_from(v).ok())
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
        .find(|abs| abs.id == num.abstract_num_id);

    let mut raw_levels: BTreeMap<u32, RawListLevel> = abstract_num
        .map(|abs| {
            abs.levels
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

fn build_numbering_map(numberings: &docx_rs::Numberings) -> NumberingMap {
    numberings
        .numberings
        .iter()
        .map(|num| (num.id, resolve_numbering(num, numberings)))
        .collect()
}

/// Extract numbering info from a paragraph, if it has numPr.
fn extract_num_info(para: &docx_rs::Paragraph) -> Option<NumInfo> {
    if !para.has_numbering {
        return None;
    }
    let np = para.property.numbering_property.as_ref()?;
    let num_id = np.id.as_ref()?.id;
    let level = np.level.as_ref().map_or(0, |l| l.val as u32);
    // numId 0 means "no numbering" in OOXML
    if num_id == 0 {
        return None;
    }
    Some(NumInfo { num_id, level })
}

/// Resolved style formatting extracted from a document style definition.
/// Contains text and paragraph formatting along with an optional heading level.
struct ResolvedStyle {
    text: TextStyle,
    paragraph: ParagraphStyle,
    paragraph_tab_overrides: Option<Vec<TabStopOverride>>,
    /// Heading level from outline_lvl (0 = Heading 1, 1 = Heading 2, ..., 5 = Heading 6).
    heading_level: Option<usize>,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum TabStopOverride {
    Set(TabStop),
    Clear(f64),
}

/// Map from style_id → resolved formatting.
type StyleMap = HashMap<String, ResolvedStyle>;

/// Synthetic style ID used for document-level default text properties.
const DOC_DEFAULT_STYLE_ID: &str = "__office2pdf_doc_defaults";

/// Default font sizes for heading levels (Heading 1-6).
/// Index 0 = Heading 1 (outline_lvl 0), index 5 = Heading 6 (outline_lvl 5).
const HEADING_DEFAULT_SIZES: [f64; 6] = [24.0, 20.0, 16.0, 14.0, 12.0, 11.0];

/// Build a map from style ID → resolved formatting by extracting formatting
/// from each style's run_property and paragraph_property.
fn build_style_map(styles: &docx_rs::Styles) -> StyleMap {
    let mut map = StyleMap::new();
    let default_text: TextStyle = extract_doc_default_text_style(styles);

    map.insert(
        DOC_DEFAULT_STYLE_ID.to_string(),
        ResolvedStyle {
            text: default_text.clone(),
            paragraph: ParagraphStyle::default(),
            paragraph_tab_overrides: None,
            heading_level: None,
        },
    );

    for style in &styles.styles {
        // Only process paragraph styles (not character or table styles)
        if style.style_type != docx_rs::StyleType::Paragraph {
            continue;
        }

        let text = merge_text_style(
            &extract_run_style(&style.run_property),
            map.get(DOC_DEFAULT_STYLE_ID),
        );
        let paragraph = extract_paragraph_style(&style.paragraph_property);
        let paragraph_tab_overrides = extract_tab_stop_overrides(&style.paragraph_property.tabs);
        let heading_level = style
            .paragraph_property
            .outline_lvl
            .as_ref()
            .map(|ol| ol.v)
            .filter(|&v| v < 6);

        map.insert(
            style.style_id.clone(),
            ResolvedStyle {
                text,
                paragraph,
                paragraph_tab_overrides,
                heading_level,
            },
        );
    }
    map
}

/// Merge style text formatting with explicit run formatting.
/// Explicit formatting (from the run itself) takes priority over style formatting.
/// For heading styles, default sizes and bold are applied when neither the style
/// nor the run specifies them.
fn merge_text_style(explicit: &TextStyle, style: Option<&ResolvedStyle>) -> TextStyle {
    let (style_text, heading_level) = match style {
        Some(s) => (&s.text, s.heading_level),
        None => return explicit.clone(),
    };

    // Start with style defaults, then apply heading defaults, then explicit overrides
    let mut merged = TextStyle {
        bold: style_text.bold,
        italic: style_text.italic,
        underline: style_text.underline,
        strikethrough: style_text.strikethrough,
        font_size: style_text.font_size,
        color: style_text.color,
        font_family: style_text.font_family.clone(),
        highlight: style_text.highlight,
        vertical_align: style_text.vertical_align,
        all_caps: style_text.all_caps,
        small_caps: style_text.small_caps,
        letter_spacing: style_text.letter_spacing,
    };

    // Apply heading defaults for missing fields
    if let Some(level) = heading_level {
        if merged.font_size.is_none() {
            merged.font_size = Some(HEADING_DEFAULT_SIZES[level]);
        }
        if merged.bold.is_none() {
            merged.bold = Some(true);
        }
    }

    // Explicit formatting overrides everything
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
fn merge_paragraph_style(
    explicit: &ParagraphStyle,
    explicit_tab_overrides: Option<&[TabStopOverride]>,
    style: Option<&ResolvedStyle>,
) -> ParagraphStyle {
    let style_para = style.map(|s| &s.paragraph);
    let inherited_tab_stops = style.and_then(resolve_style_tab_stops);

    ParagraphStyle {
        alignment: explicit.alignment.or(style_para.and_then(|s| s.alignment)),
        indent_left: explicit
            .indent_left
            .or(style_para.and_then(|s| s.indent_left)),
        indent_right: explicit
            .indent_right
            .or(style_para.and_then(|s| s.indent_right)),
        indent_first_line: explicit
            .indent_first_line
            .or(style_para.and_then(|s| s.indent_first_line)),
        line_spacing: explicit
            .line_spacing
            .or(style_para.and_then(|s| s.line_spacing)),
        space_before: explicit
            .space_before
            .or(style_para.and_then(|s| s.space_before)),
        space_after: explicit
            .space_after
            .or(style_para.and_then(|s| s.space_after)),
        // Heading level from the style definition (outline_lvl 0→H1, 1→H2, ...)
        heading_level: style
            .and_then(|s| s.heading_level)
            .map(|lvl| (lvl + 1) as u8),
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
        .map(|explicit_tab_stops| explicit_tab_stops.to_vec())
        .or_else(|| inherited_tab_stops.map(|inherited_tab_stops| inherited_tab_stops.to_vec()))
}

fn apply_tab_stop_overrides(tab_stops: &mut Vec<TabStop>, tab_overrides: &[TabStopOverride]) {
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
fn get_paragraph_style_id(prop: &docx_rs::ParagraphProperty) -> Option<&str> {
    prop.style.as_ref().map(|s| s.val.as_str())
}

/// An intermediate element that carries optional numbering info alongside blocks.
enum TaggedElement {
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
            (Some(level), Some(prev_level))
                if level.style.kind == ListKind::Ordered && item.level > prev_level =>
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
fn group_into_lists(elements: Vec<TaggedElement>, numberings: &NumberingMap) -> Vec<Block> {
    let mut result: Vec<Block> = Vec::new();

    // Accumulator for current list run
    let mut current_list: Option<(usize, Vec<ListItem>)> = None; // (numId, items)

    for elem in elements {
        match elem {
            TaggedElement::ListParagraph { info, paragraph } => {
                if let Some((cur_num_id, ref mut items)) = current_list {
                    if info.num_id == cur_num_id {
                        // Same list — add item
                        items.push(ListItem {
                            content: vec![paragraph],
                            level: info.level,
                            start_at: None,
                        });
                        continue;
                    }
                    // Different list — flush current
                    result.push(Block::List(finalize_list(
                        cur_num_id,
                        std::mem::take(items),
                        numberings,
                    )));
                }
                // Start new list
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
                // Flush any pending list
                if let Some((num_id, items)) = current_list.take() {
                    result.push(Block::List(finalize_list(num_id, items, numberings)));
                }
                result.extend(blocks);
            }
        }
    }

    // Flush trailing list
    if let Some((num_id, items)) = current_list {
        result.push(Block::List(finalize_list(num_id, items, numberings)));
    }

    result
}

// ── Footnote / Endnote support ──────────────────────────────────────────

/// A note reference kind.
#[derive(Debug, Clone, Copy)]
enum NoteKind {
    Footnote,
    Endnote,
}

/// Context for resolving footnote/endnote references during parsing.
/// The `cursor` is advanced each time a note reference run is encountered.
struct NoteContext {
    /// Footnote ID → plain text content.
    footnote_content: HashMap<usize, String>,
    /// Endnote ID → plain text content.
    endnote_content: HashMap<usize, String>,
    /// Ordered list of note references as they appear in document.xml.
    note_refs: Vec<(NoteKind, usize)>,
    /// Current position in `note_refs`.
    cursor: Cell<usize>,
    /// Style IDs that indicate footnote/endnote reference runs.
    /// Includes English defaults plus any locale-specific IDs found in styles.
    note_style_ids: HashSet<String>,
}

impl NoteContext {
    /// Consume the next note reference and return its text content.
    fn consume_next(&self) -> Option<String> {
        let idx = self.cursor.get();
        if idx >= self.note_refs.len() {
            return None;
        }
        let (kind, id) = self.note_refs[idx];
        self.cursor.set(idx + 1);
        match kind {
            NoteKind::Footnote => self.footnote_content.get(&id).cloned(),
            NoteKind::Endnote => self.endnote_content.get(&id).cloned(),
        }
    }

    /// Populate note style IDs from docx styles.
    /// Scans character styles whose canonical name (w:name) is "footnote reference"
    /// or "endnote reference", handling localized style IDs (e.g., German "Funotenzeichen").
    fn populate_style_ids(&mut self, styles: &docx_rs::Styles) {
        for style in &styles.styles {
            if let Ok(name_val) = serde_json::to_value(&style.name)
                && let Some(name_str) = name_val.as_str()
            {
                let lower = name_str.to_lowercase();
                if lower == "footnote reference" || lower == "endnote reference" {
                    self.note_style_ids.insert(style.style_id.clone());
                }
            }
        }
    }
}

/// Wrap type info for an anchored drawing, scanned from raw document XML.
struct AnchorWrapInfo {
    wrap_mode: WrapMode,
    behind_doc: bool,
}

/// Context for resolving wrap modes of anchored drawings during parsing.
/// The `cursor` is advanced each time an anchored drawing is encountered.
struct WrapContext {
    /// Ordered list of wrap info for anchored drawings as they appear in document.xml.
    wraps: Vec<AnchorWrapInfo>,
    /// Current position in `wraps`.
    cursor: Cell<usize>,
}

impl WrapContext {
    /// Consume the next anchor wrap info and return its wrap mode.
    fn consume_next(&self) -> WrapMode {
        let idx = self.cursor.get();
        if idx >= self.wraps.len() {
            return WrapMode::None;
        }
        let info = &self.wraps[idx];
        self.cursor.set(idx + 1);
        // behindDoc attribute overrides to Behind mode
        if info.behind_doc {
            WrapMode::Behind
        } else {
            info.wrap_mode
        }
    }
}

/// Build a `WrapContext` from a pre-read document.xml string.
fn build_wrap_context_from_xml(doc_xml: Option<&str>) -> WrapContext {
    let wraps = doc_xml.map(scan_anchor_wrap_types).unwrap_or_default();
    WrapContext {
        wraps,
        cursor: Cell::new(0),
    }
}

#[derive(Debug, Clone, Copy, Default)]
struct DrawingTextBoxInfo {
    width_pt: Option<f64>,
    height_pt: Option<f64>,
}

/// Context for resolving DrawingML text box extents in document order.
struct DrawingTextBoxContext {
    text_boxes: Vec<DrawingTextBoxInfo>,
    cursor: Cell<usize>,
}

impl DrawingTextBoxContext {
    fn from_xml(xml: Option<&str>) -> Self {
        Self {
            text_boxes: xml.map(scan_drawing_text_boxes).unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    fn consume_next(&self) -> DrawingTextBoxInfo {
        let idx: usize = self.cursor.get();
        self.cursor.set(idx + 1);
        self.text_boxes.get(idx).copied().unwrap_or_default()
    }
}

#[derive(Debug, Clone, Copy, Default)]
struct TableHeaderInfo {
    repeat_rows: usize,
}

/// Context for resolving repeat-header rows on tables.
/// docx-rs does not expose `w:tblHeader`, so we scan raw XML and consume the
/// results in table encounter order.
struct TableHeaderContext {
    headers: Vec<TableHeaderInfo>,
    cursor: Cell<usize>,
}

impl TableHeaderContext {
    fn from_xml(xml: Option<&str>) -> Self {
        Self {
            headers: xml.map(scan_table_headers).unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    fn consume_next(&self) -> TableHeaderInfo {
        let idx: usize = self.cursor.get();
        self.cursor.set(idx + 1);
        self.headers.get(idx).copied().unwrap_or_default()
    }
}

struct TableHeaderScanState {
    table_index: usize,
    repeat_rows: usize,
    in_row: bool,
    current_row_is_header: bool,
    saw_body_row: bool,
}

fn scan_table_headers(xml: &str) -> Vec<TableHeaderInfo> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf: Vec<u8> = Vec::new();
    let mut headers: Vec<TableHeaderInfo> = Vec::new();
    let mut stack: Vec<TableHeaderScanState> = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => match e.local_name().as_ref() {
                b"tbl" => {
                    headers.push(TableHeaderInfo::default());
                    stack.push(TableHeaderScanState {
                        table_index: headers.len() - 1,
                        repeat_rows: 0,
                        in_row: false,
                        current_row_is_header: false,
                        saw_body_row: false,
                    });
                }
                b"tr" => {
                    if let Some(state) = stack.last_mut() {
                        state.in_row = true;
                        state.current_row_is_header = false;
                    }
                }
                b"tblHeader" => {
                    if let Some(state) = stack.last_mut()
                        && state.in_row
                        && on_off_element_is_enabled(e)
                    {
                        state.current_row_is_header = true;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"tbl" => {
                    headers.push(TableHeaderInfo::default());
                }
                b"tr" => {
                    if let Some(state) = stack.last_mut() {
                        state.in_row = true;
                        state.current_row_is_header = false;
                        finalize_table_header_row(state);
                    }
                }
                b"tblHeader" => {
                    if let Some(state) = stack.last_mut()
                        && state.in_row
                        && on_off_element_is_enabled(e)
                    {
                        state.current_row_is_header = true;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::End(ref e)) => match e.local_name().as_ref() {
                b"tr" => {
                    if let Some(state) = stack.last_mut() {
                        finalize_table_header_row(state);
                    }
                }
                b"tbl" => {
                    if let Some(state) = stack.pop() {
                        headers[state.table_index].repeat_rows = state.repeat_rows;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    headers
}

fn finalize_table_header_row(state: &mut TableHeaderScanState) {
    if !state.in_row {
        return;
    }

    if !state.saw_body_row && state.current_row_is_header {
        state.repeat_rows += 1;
    } else {
        state.saw_body_row = true;
    }

    state.in_row = false;
    state.current_row_is_header = false;
}

fn on_off_element_is_enabled(e: &quick_xml::events::BytesStart<'_>) -> bool {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() != b"val" {
            continue;
        }

        let value = attr.value.as_ref();
        if value.eq_ignore_ascii_case(b"0")
            || value.eq_ignore_ascii_case(b"false")
            || value.eq_ignore_ascii_case(b"off")
        {
            return false;
        }
    }

    true
}

fn scan_drawing_text_boxes(xml: &str) -> Vec<DrawingTextBoxInfo> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf: Vec<u8> = Vec::new();
    let mut result: Vec<DrawingTextBoxInfo> = Vec::new();
    let mut in_body: bool = false;
    let mut drawing_depth: usize = 0;
    let mut current_info: DrawingTextBoxInfo = DrawingTextBoxInfo::default();
    let mut saw_text_box: bool = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => match e.local_name().as_ref() {
                b"body" => in_body = true,
                b"drawing" if in_body => {
                    if drawing_depth == 0 {
                        current_info = DrawingTextBoxInfo::default();
                        saw_text_box = false;
                    }
                    drawing_depth += 1;
                }
                b"extent" if drawing_depth > 0 => {
                    update_drawing_text_box_extent(&mut current_info, e);
                }
                b"txbx" if drawing_depth > 0 => saw_text_box = true,
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"extent" if drawing_depth > 0 => {
                    update_drawing_text_box_extent(&mut current_info, e);
                }
                b"txbx" if drawing_depth > 0 => saw_text_box = true,
                _ => {}
            },
            Ok(quick_xml::events::Event::End(ref e)) => match e.local_name().as_ref() {
                b"body" => in_body = false,
                b"drawing" if drawing_depth > 0 => {
                    drawing_depth -= 1;
                    if drawing_depth == 0 && saw_text_box {
                        result.push(current_info);
                        current_info = DrawingTextBoxInfo::default();
                        saw_text_box = false;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    result
}

fn update_drawing_text_box_extent(
    info: &mut DrawingTextBoxInfo,
    e: &quick_xml::events::BytesStart<'_>,
) {
    if info.width_pt.is_some() && info.height_pt.is_some() {
        return;
    }

    let mut width_emu: Option<u32> = None;
    let mut height_emu: Option<u32> = None;

    for attr in e.attributes().flatten() {
        match attr.key.local_name().as_ref() {
            b"cx" => {
                width_emu = std::str::from_utf8(attr.value.as_ref())
                    .ok()
                    .and_then(|value| value.parse::<u32>().ok());
            }
            b"cy" => {
                height_emu = std::str::from_utf8(attr.value.as_ref())
                    .ok()
                    .and_then(|value| value.parse::<u32>().ok());
            }
            _ => {}
        }
    }

    if let Some(width_emu) = width_emu {
        info.width_pt = Some(emu_to_pt(width_emu));
    }
    if let Some(height_emu) = height_emu {
        info.height_pt = Some(emu_to_pt(height_emu));
    }
}

#[derive(Debug, Clone, Default)]
struct VmlTextBoxInfo {
    paragraphs: Vec<String>,
    wrap_mode: Option<WrapMode>,
}

impl VmlTextBoxInfo {
    fn into_blocks(self) -> Vec<Block> {
        self.paragraphs
            .into_iter()
            .filter(|text| !text.is_empty())
            .map(|text| {
                Block::Paragraph(Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text,
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                })
            })
            .collect()
    }
}

/// Raw VML shape textbox content scanned in body order.
struct VmlTextBoxContext {
    text_boxes: Vec<VmlTextBoxInfo>,
    cursor: Cell<usize>,
}

impl VmlTextBoxContext {
    fn from_xml(xml: Option<&str>) -> Self {
        Self {
            text_boxes: xml.map(scan_vml_text_boxes).unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    fn consume_next(&self) -> VmlTextBoxInfo {
        let idx: usize = self.cursor.get();
        self.cursor.set(idx + 1);
        self.text_boxes.get(idx).cloned().unwrap_or_default()
    }
}

fn scan_vml_text_boxes(xml: &str) -> Vec<VmlTextBoxInfo> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buf: Vec<u8> = Vec::new();
    let mut result: Vec<VmlTextBoxInfo> = Vec::new();
    let mut in_body: bool = false;
    let mut pict_depth: usize = 0;
    let mut shape_depth: usize = 0;
    let mut in_txbx_content: bool = false;
    let mut in_paragraph: bool = false;
    let mut current_pict_shapes: Vec<VmlTextBoxInfo> = Vec::new();
    let mut current_pict_wrap: Option<WrapMode> = None;
    let mut current_shape_paragraphs: Vec<String> = Vec::new();
    let mut current_paragraph_text: String = String::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(quick_xml::events::Event::Start(ref e)) => match e.local_name().as_ref() {
                b"body" => in_body = true,
                b"pict" if in_body => {
                    if pict_depth == 0 {
                        current_pict_shapes.clear();
                        current_pict_wrap = None;
                    }
                    pict_depth += 1;
                }
                b"shape" if pict_depth > 0 => {
                    if shape_depth == 0 {
                        current_shape_paragraphs.clear();
                    }
                    shape_depth += 1;
                }
                b"txbxContent" if shape_depth > 0 => in_txbx_content = true,
                b"p" if in_txbx_content => {
                    in_paragraph = true;
                    current_paragraph_text.clear();
                }
                b"wrap" if pict_depth > 0 => {
                    current_pict_wrap = extract_vml_wrap_mode_from_element(e);
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"tab" if in_paragraph => current_paragraph_text.push('\t'),
                b"br" if in_paragraph => current_paragraph_text.push('\n'),
                b"wrap" if pict_depth > 0 => {
                    current_pict_wrap = extract_vml_wrap_mode_from_element(e);
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Text(ref e)) => {
                if in_paragraph && let Ok(text) = e.xml_content() {
                    current_paragraph_text.push_str(&text);
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => match e.local_name().as_ref() {
                b"body" => in_body = false,
                b"p" if in_paragraph => {
                    current_shape_paragraphs.push(std::mem::take(&mut current_paragraph_text));
                    in_paragraph = false;
                }
                b"txbxContent" if in_txbx_content => in_txbx_content = false,
                b"shape" if shape_depth > 0 => {
                    shape_depth -= 1;
                    if shape_depth == 0 {
                        current_pict_shapes.push(VmlTextBoxInfo {
                            paragraphs: std::mem::take(&mut current_shape_paragraphs),
                            wrap_mode: None,
                        });
                        in_txbx_content = false;
                        in_paragraph = false;
                        current_paragraph_text.clear();
                    }
                }
                b"pict" if pict_depth > 0 => {
                    pict_depth -= 1;
                    if pict_depth == 0 {
                        for mut text_box in current_pict_shapes.drain(..) {
                            text_box.wrap_mode = current_pict_wrap;
                            result.push(text_box);
                        }
                        current_pict_wrap = None;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    result
}

fn extract_vml_wrap_mode_from_element(e: &quick_xml::events::BytesStart<'_>) -> Option<WrapMode> {
    for attr in e.attributes().flatten() {
        if attr.key.local_name().as_ref() != b"type" {
            continue;
        }

        let value = std::str::from_utf8(attr.value.as_ref()).ok()?;
        return match value {
            "square" => Some(WrapMode::Square),
            "none" => Some(WrapMode::None),
            "tight" | "through" => Some(WrapMode::Tight),
            "topAndBottom" | "top-and-bottom" => Some(WrapMode::TopAndBottom),
            _ => None,
        };
    }

    None
}

/// Scan document.xml for `<wp:anchor>` elements and extract their wrap type.
/// Returns wrap info in document order for correlation with docx-rs parsed images.
fn scan_anchor_wrap_types(xml: &str) -> Vec<AnchorWrapInfo> {
    let mut results = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    let mut in_anchor = false;
    let mut behind_doc = false;
    let mut found_wrap = false;
    let mut current_wrap = WrapMode::None;

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();
                match name {
                    b"anchor" => {
                        in_anchor = true;
                        behind_doc = false;
                        found_wrap = false;
                        current_wrap = WrapMode::None;
                        // Check for behindDoc attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"behindDoc"
                                && let Ok(val) = attr.unescape_value()
                                && (val == "1" || val == "true")
                            {
                                behind_doc = true;
                            }
                        }
                    }
                    b"wrapSquare" if in_anchor => {
                        current_wrap = WrapMode::Square;
                        found_wrap = true;
                    }
                    b"wrapTight" if in_anchor => {
                        current_wrap = WrapMode::Tight;
                        found_wrap = true;
                    }
                    b"wrapTopAndBottom" if in_anchor => {
                        current_wrap = WrapMode::TopAndBottom;
                        found_wrap = true;
                    }
                    b"wrapNone" if in_anchor => {
                        current_wrap = WrapMode::None;
                        found_wrap = true;
                    }
                    b"wrapThrough" if in_anchor => {
                        // Treat wrapThrough similar to Tight
                        current_wrap = WrapMode::Tight;
                        found_wrap = true;
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"anchor" && in_anchor {
                    // If no explicit wrap element was found and behindDoc is set,
                    // treat as Behind. Otherwise default to None.
                    if !found_wrap && behind_doc {
                        current_wrap = WrapMode::None; // behind_doc flag handled in WrapContext
                    }
                    results.push(AnchorWrapInfo {
                        wrap_mode: current_wrap,
                        behind_doc,
                    });
                    in_anchor = false;
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    results
}

/// Context for tracking bidi (right-to-left) paragraphs.
/// Built by scanning the raw document XML for `<w:bidi/>` elements within `<w:pPr>`.
/// docx-rs does not read `w:bidi` back during parsing, so raw XML scanning is needed.
struct BidiContext {
    /// Set of 0-based paragraph indices (ALL `<w:p>` in document order) that are bidi.
    bidi_indices: HashSet<usize>,
    /// Auto-incrementing counter consumed by the paragraph converter.
    cursor: Cell<usize>,
}

impl BidiContext {
    fn from_xml(xml: Option<&str>) -> Self {
        let bidi_indices = xml.map(Self::scan).unwrap_or_default();
        Self {
            bidi_indices,
            cursor: Cell::new(0),
        }
    }

    /// Advance the cursor and return whether the current paragraph is bidi.
    fn next_is_bidi(&self) -> bool {
        let idx = self.cursor.get();
        self.cursor.set(idx + 1);
        self.bidi_indices.contains(&idx)
    }

    fn scan(xml: &str) -> HashSet<usize> {
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut buf = Vec::new();
        let mut result = HashSet::new();
        let mut para_index: usize = 0;
        let mut in_ppr = false;
        let mut in_body = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(
                    quick_xml::events::Event::Start(ref e) | quick_xml::events::Event::Empty(ref e),
                ) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"body" => in_body = true,
                        b"pPr" if in_body => in_ppr = true,
                        b"bidi" if in_ppr => {
                            result.insert(para_index);
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"body" => in_body = false,
                        b"p" if in_body => {
                            para_index += 1;
                            in_ppr = false;
                        }
                        b"pPr" => in_ppr = false,
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
        result
    }
}

/// Pre-scanned context for `<w:smallCaps>` in run properties.
/// docx-rs does not expose the `smallCaps` element, so we scan the raw XML.
/// Tracks per-run smallCaps flags ordered by document body appearance.
struct SmallCapsContext {
    /// Flat list of booleans, one per `<w:r>` encountered in document body order.
    flags: Vec<bool>,
    /// Cursor for consuming flags in order during conversion.
    cursor: Cell<usize>,
}

impl SmallCapsContext {
    fn from_xml(xml: Option<&str>) -> Self {
        let flags: Vec<bool> = xml.map(Self::scan).unwrap_or_default();
        Self {
            flags,
            cursor: Cell::new(0),
        }
    }

    /// Advance the cursor and return whether the current run has smallCaps.
    fn next_is_small_caps(&self) -> bool {
        let idx: usize = self.cursor.get();
        self.cursor.set(idx + 1);
        self.flags.get(idx).copied().unwrap_or(false)
    }

    fn scan(xml: &str) -> Vec<bool> {
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut buf: Vec<u8> = Vec::new();
        let mut result: Vec<bool> = Vec::new();
        let mut in_body: bool = false;
        let mut in_run: bool = false;
        let mut in_rpr: bool = false;
        let mut current_has_small_caps: bool = false;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(
                    quick_xml::events::Event::Start(ref e) | quick_xml::events::Event::Empty(ref e),
                ) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"body" => in_body = true,
                        b"r" if in_body => {
                            in_run = true;
                            current_has_small_caps = false;
                        }
                        b"rPr" if in_run => in_rpr = true,
                        b"smallCaps" if in_rpr => {
                            // Check for w:val="false" / w:val="0" to handle explicit disable
                            let is_disabled: bool = e.attributes().flatten().any(|a| {
                                a.key.local_name().as_ref() == b"val"
                                    && matches!(a.value.as_ref(), b"false" | b"0")
                            });
                            if !is_disabled {
                                current_has_small_caps = true;
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::End(ref e)) => {
                    let local = e.local_name();
                    match local.as_ref() {
                        b"body" => in_body = false,
                        b"r" if in_body => {
                            result.push(current_has_small_caps);
                            in_run = false;
                            in_rpr = false;
                            current_has_small_caps = false;
                        }
                        b"rPr" => in_rpr = false,
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buf.clear();
        }
        result
    }
}

/// Scan document.xml for `<w:cols>` within each `<w:sectPr>` in document order.
/// docx-rs exposes equal-width column count and spacing, but not unequal widths.
fn scan_column_layouts(xml: &str) -> Vec<Option<ColumnLayout>> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut layouts: Vec<Option<ColumnLayout>> = Vec::new();

    let mut in_sect_pr = false;
    let mut in_cols = false;
    let mut num_columns: u32 = 1;
    let mut spacing_twips: f64 = 720.0; // default 720 twips = 36pt
    let mut equal_width = true;
    let mut col_widths: Vec<f64> = Vec::new();

    let build_layout =
        |num_columns: u32, spacing_twips: f64, equal_width: bool, col_widths: &[f64]| {
            if num_columns < 2 {
                return None;
            }

            let column_widths = if !equal_width && !col_widths.is_empty() {
                Some(col_widths.to_vec())
            } else {
                None
            };

            Some(ColumnLayout {
                num_columns,
                spacing: spacing_twips / 20.0, // twips → points
                column_widths,
            })
        };

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();
                match name {
                    b"sectPr" => {
                        in_sect_pr = true;
                        // Reset for each sectPr.
                        num_columns = 1;
                        spacing_twips = 720.0;
                        equal_width = true;
                        col_widths.clear();
                    }
                    b"cols" if in_sect_pr => {
                        in_cols = true;
                        for attr in e.attributes().flatten() {
                            let key = attr.key.local_name();
                            if let Ok(val) = attr.unescape_value() {
                                match key.as_ref() {
                                    b"num" => {
                                        if let Ok(n) = val.parse::<u32>() {
                                            num_columns = n;
                                        }
                                    }
                                    b"space" => {
                                        if let Ok(s) = val.parse::<f64>() {
                                            spacing_twips = s;
                                        }
                                    }
                                    b"equalWidth" => {
                                        equal_width = val != "0";
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    b"col" if in_cols => {
                        // Per-column width specification
                        for attr in e.attributes().flatten() {
                            let key = attr.key.local_name();
                            if key.as_ref() == b"w"
                                && let Ok(val) = attr.unescape_value()
                                && let Ok(w) = val.parse::<f64>()
                            {
                                col_widths.push(w / 20.0); // twips → points
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();
                match name {
                    b"sectPr" => {
                        layouts.push(build_layout(1, 720.0, true, &[]));
                    }
                    b"cols" if in_sect_pr => {
                        in_cols = false;
                        for attr in e.attributes().flatten() {
                            let key = attr.key.local_name();
                            if let Ok(val) = attr.unescape_value() {
                                match key.as_ref() {
                                    b"num" => {
                                        if let Ok(n) = val.parse::<u32>() {
                                            num_columns = n;
                                        }
                                    }
                                    b"space" => {
                                        if let Ok(s) = val.parse::<f64>() {
                                            spacing_twips = s;
                                        }
                                    }
                                    b"equalWidth" => {
                                        equal_width = val != "0";
                                    }
                                    _ => {}
                                }
                            }
                        }
                    }
                    b"col" if in_cols => {
                        for attr in e.attributes().flatten() {
                            let key = attr.key.local_name();
                            if key.as_ref() == b"w"
                                && let Ok(val) = attr.unescape_value()
                                && let Ok(w) = val.parse::<f64>()
                            {
                                col_widths.push(w / 20.0); // twips → points
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"sectPr" => {
                        layouts.push(build_layout(
                            num_columns,
                            spacing_twips,
                            equal_width,
                            &col_widths,
                        ));
                        in_sect_pr = false;
                    }
                    b"cols" => in_cols = false,
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    layouts
}

fn extract_column_layout_from_section_property(
    section_prop: &docx_rs::SectionProperty,
) -> Option<ColumnLayout> {
    if section_prop.columns < 2 {
        return None;
    }

    Some(ColumnLayout {
        num_columns: section_prop.columns as u32,
        spacing: section_prop.space as f64 / 20.0,
        column_widths: None,
    })
}

/// Context for OMML math equations extracted from raw document XML.
/// docx-rs does not parse `m:oMath` / `m:oMathPara` elements — they are
/// completely absent from the `ParagraphChild` enum — so we scan the raw ZIP.
struct MathContext {
    /// Math equations keyed by 0-based body child index in document.xml.
    /// A paragraph can contain multiple equations.
    equations: HashMap<usize, Vec<MathEquation>>,
}

impl MathContext {
    /// Take the math equations for a given body child index (consuming them).
    fn take(&mut self, index: usize) -> Vec<MathEquation> {
        self.equations.remove(&index).unwrap_or_default()
    }
}

/// Build a `MathContext` from a pre-read document.xml string.
fn build_math_context_from_xml(doc_xml: Option<&str>) -> MathContext {
    let mut equations: HashMap<usize, Vec<MathEquation>> = HashMap::new();

    if let Some(xml) = doc_xml {
        let raw = super::omml::scan_math_equations(xml);
        for (idx, content, display) in raw {
            equations
                .entry(idx)
                .or_default()
                .push(MathEquation { content, display });
        }
    }

    MathContext { equations }
}

/// Context for embedded charts extracted from raw DOCX ZIP.
/// docx-rs does not parse chart drawing elements, so we scan the raw ZIP.
struct ChartContext {
    /// Charts keyed by 0-based body child index in document.xml.
    charts: HashMap<usize, Vec<Chart>>,
}

impl ChartContext {
    /// Take the charts for a given body child index (consuming them).
    fn take(&mut self, index: usize) -> Vec<Chart> {
        self.charts.remove(&index).unwrap_or_default()
    }
}

/// Build a `ChartContext` from pre-read XML strings and a shared ZIP archive.
fn build_chart_context_from_xml(
    doc_xml: Option<&str>,
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
) -> ChartContext {
    let mut charts: HashMap<usize, Vec<Chart>> = HashMap::new();

    let Some(doc_xml) = doc_xml else {
        return ChartContext { charts };
    };

    let rels_xml = read_zip_text(archive, "word/_rels/document.xml.rels");
    let Some(rels_xml) = rels_xml else {
        return ChartContext { charts };
    };

    // Find chart references in document.xml
    let refs = super::chart::scan_chart_references(doc_xml);
    // Build relationship ID → chart file path mapping
    let rels = super::chart::scan_chart_rels(&rels_xml);

    // For each chart reference, parse the chart XML
    for (body_idx, rid) in refs {
        if let Some(chart_path) = rels.get(&rid)
            && let Some(chart_xml) = read_zip_text(archive, chart_path)
            && let Some(chart) = super::chart::parse_chart_xml(&chart_xml)
        {
            charts.entry(body_idx).or_default().push(chart);
        }
    }

    ChartContext { charts }
}

/// Build a `NoteContext` from pre-read XML strings and a shared ZIP archive.
fn build_note_context_from_xml(
    doc_xml: Option<&str>,
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
) -> NoteContext {
    let default_style_ids: HashSet<String> = ["FootnoteReference", "EndnoteReference"]
        .iter()
        .map(|s| (*s).to_string())
        .collect();

    let mut footnote_content = HashMap::new();
    let mut endnote_content = HashMap::new();

    // Parse word/footnotes.xml
    if let Some(xml) = read_zip_text(archive, "word/footnotes.xml") {
        footnote_content = parse_notes_xml(&xml);
    }

    // Parse word/endnotes.xml
    if let Some(xml) = read_zip_text(archive, "word/endnotes.xml") {
        endnote_content = parse_notes_xml(&xml);
    }

    // Scan word/document.xml for note references in document order
    let note_refs = doc_xml.map(scan_note_refs).unwrap_or_default();

    NoteContext {
        footnote_content,
        endnote_content,
        note_refs,
        cursor: Cell::new(0),
        note_style_ids: default_style_ids,
    }
}

/// Read a ZIP entry as a UTF-8 string.
fn read_zip_text(archive: &mut zip::ZipArchive<impl Read + Seek>, name: &str) -> Option<String> {
    let mut file = archive.by_name(name).ok()?;
    let mut contents = String::new();
    file.read_to_string(&mut contents).ok()?;
    Some(contents)
}

/// Parse footnotes or endnotes XML into a map of ID → concatenated text.
/// Works for both `<w:footnote w:id="N">` and `<w:endnote w:id="N">` elements.
fn parse_notes_xml(xml: &str) -> HashMap<usize, String> {
    let mut map = HashMap::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut current_id: Option<usize> = None;
    let mut current_text = String::new();
    let mut in_text = false;

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();
                match name {
                    b"footnote" | b"endnote" => {
                        // Save previous note if any
                        if let Some(id) = current_id.take() {
                            let text = current_text.trim().to_string();
                            if !text.is_empty() {
                                map.insert(id, text);
                            }
                        }
                        current_text.clear();
                        // Extract w:id attribute
                        for attr in e.attributes().flatten() {
                            if attr.key.local_name().as_ref() == b"id"
                                && let Ok(val) = attr.unescape_value()
                            {
                                current_id = val.parse::<usize>().ok();
                            }
                        }
                    }
                    b"t" => in_text = true,
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();
                match name {
                    b"t" => in_text = false,
                    b"footnote" | b"endnote" => {
                        if let Some(id) = current_id.take() {
                            let text = current_text.trim().to_string();
                            if !text.is_empty() {
                                map.insert(id, text);
                            }
                        }
                        current_text.clear();
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::Text(ref e)) => {
                if in_text && let Ok(text) = e.xml_content() {
                    if !current_text.is_empty() {
                        current_text.push(' ');
                    }
                    current_text.push_str(&text);
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    map
}

/// Scan document.xml for `<w:footnoteReference>` and `<w:endnoteReference>` elements,
/// returning them in document order with their IDs.
fn scan_note_refs(xml: &str) -> Vec<(NoteKind, usize)> {
    let mut refs = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref e))
            | Ok(quick_xml::events::Event::Empty(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();
                let kind = match name {
                    b"footnoteReference" => Some(NoteKind::Footnote),
                    b"endnoteReference" => Some(NoteKind::Endnote),
                    _ => None,
                };
                if let Some(kind) = kind {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"id"
                            && let Ok(val) = attr.unescape_value()
                            && let Ok(id) = val.parse::<usize>()
                        {
                            refs.push((kind, id));
                        }
                    }
                }
            }
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    refs
}

/// Check if a docx-rs Run represents a footnote or endnote reference.
/// These runs have `rStyle` set to a footnote/endnote reference style
/// (e.g., "FootnoteReference" in English, "Funotenzeichen" in German)
/// and contain no text.
fn is_note_reference_run(run: &docx_rs::Run, notes: &NoteContext) -> bool {
    if let Some(ref style) = run.run_property.style
        && notes.note_style_ids.contains(&style.val)
    {
        // Verify the run has no text content (only footnoteReference element)
        return extract_run_text(run).is_empty();
    }
    false
}

fn build_flow_page_from_section(
    section_prop: &docx_rs::SectionProperty,
    elements: Vec<TaggedElement>,
    numberings: &NumberingMap,
    header_footer_assets: &HeaderFooterAssets,
    column_layout: Option<ColumnLayout>,
    warnings: &mut Vec<ConvertWarning>,
) -> FlowPage {
    let (size, margins) = extract_page_setup(section_prop);
    let content = group_into_lists(elements, numberings);

    for block in &content {
        if let Block::Chart(chart) = block {
            let title = chart.title.as_deref().unwrap_or("untitled").to_string();
            warnings.push(ConvertWarning::FallbackUsed {
                format: "DOCX".to_string(),
                from: format!("chart ({title})"),
                to: "data table".to_string(),
            });
        }
    }

    if matches!(
        section_prop.section_type,
        Some(docx_rs::SectionType::Continuous | docx_rs::SectionType::NextColumn)
    ) {
        warnings.push(ConvertWarning::FallbackUsed {
            format: "DOCX".to_string(),
            from: "continuous section break".to_string(),
            to: "page-level section split".to_string(),
        });
    }

    if section_prop.first_header_reference.is_some()
        || section_prop.first_footer_reference.is_some()
        || section_prop.even_header_reference.is_some()
        || section_prop.even_footer_reference.is_some()
        || section_prop.first_header.is_some()
        || section_prop.first_footer.is_some()
        || section_prop.even_header.is_some()
        || section_prop.even_footer.is_some()
    {
        warnings.push(ConvertWarning::FallbackUsed {
            format: "DOCX".to_string(),
            from: "header/footer variants".to_string(),
            to: "single header/footer per section".to_string(),
        });
    }

    if section_prop
        .page_num_type
        .as_ref()
        .and_then(|page_num_type| page_num_type.start)
        .is_some()
    {
        warnings.push(ConvertWarning::FallbackUsed {
            format: "DOCX".to_string(),
            from: "section page number restart".to_string(),
            to: "global page counter".to_string(),
        });
    }

    FlowPage {
        size,
        margins,
        content,
        header: extract_docx_header(section_prop, header_footer_assets),
        footer: extract_docx_footer(section_prop, header_footer_assets),
        columns: column_layout
            .or_else(|| extract_column_layout_from_section_property(section_prop)),
    }
}

impl Parser for DocxParser {
    fn parse(
        &self,
        data: &[u8],
        _options: &ConvertOptions,
    ) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
        // Open ZIP once and build all pre-parse contexts from a single pass.
        // This consolidates what was previously 5 separate ZIP opens + multiple
        // reads of word/document.xml into a single archive + single doc read.
        let (
            metadata,
            mut notes,
            wraps,
            drawing_text_boxes,
            table_headers,
            vml_text_boxes,
            mut math,
            mut chart_ctx,
            column_layouts,
            bidi,
            small_caps,
            header_footer_assets,
        ) = {
            let cursor = std::io::Cursor::new(data);
            match zip::ZipArchive::new(cursor) {
                Ok(mut archive) => {
                    let metadata = crate::parser::metadata::extract_metadata_from_zip(&mut archive);
                    let doc_xml = read_zip_text(&mut archive, "word/document.xml");
                    let notes = build_note_context_from_xml(doc_xml.as_deref(), &mut archive);
                    let wraps = build_wrap_context_from_xml(doc_xml.as_deref());
                    let drawing_text_boxes = DrawingTextBoxContext::from_xml(doc_xml.as_deref());
                    let table_headers = TableHeaderContext::from_xml(doc_xml.as_deref());
                    let vml_text_boxes = VmlTextBoxContext::from_xml(doc_xml.as_deref());
                    let math = build_math_context_from_xml(doc_xml.as_deref());
                    let chart_ctx = build_chart_context_from_xml(doc_xml.as_deref(), &mut archive);
                    let column_layouts = doc_xml
                        .as_deref()
                        .map(scan_column_layouts)
                        .unwrap_or_default();
                    let bidi = BidiContext::from_xml(doc_xml.as_deref());
                    let small_caps = SmallCapsContext::from_xml(doc_xml.as_deref());
                    let header_footer_assets = build_header_footer_assets(&mut archive);
                    (
                        metadata,
                        notes,
                        wraps,
                        drawing_text_boxes,
                        table_headers,
                        vml_text_boxes,
                        math,
                        chart_ctx,
                        column_layouts,
                        bidi,
                        small_caps,
                        header_footer_assets,
                    )
                }
                Err(_) => {
                    // ZIP open failed — return empty contexts; docx-rs will
                    // produce a proper parse error downstream.
                    let default_style_ids: HashSet<String> =
                        ["FootnoteReference", "EndnoteReference"]
                            .iter()
                            .map(|s| (*s).to_string())
                            .collect();
                    (
                        crate::ir::Metadata::default(),
                        NoteContext {
                            footnote_content: HashMap::new(),
                            endnote_content: HashMap::new(),
                            note_refs: Vec::new(),
                            cursor: Cell::new(0),
                            note_style_ids: default_style_ids,
                        },
                        WrapContext {
                            wraps: Vec::new(),
                            cursor: Cell::new(0),
                        },
                        DrawingTextBoxContext::from_xml(None),
                        TableHeaderContext::from_xml(None),
                        VmlTextBoxContext::from_xml(None),
                        MathContext {
                            equations: HashMap::new(),
                        },
                        ChartContext {
                            charts: HashMap::new(),
                        },
                        Vec::new(),
                        BidiContext::from_xml(None),
                        SmallCapsContext::from_xml(None),
                        HeaderFooterAssets::default(),
                    )
                }
            }
        };

        let docx = docx_rs::read_docx(data)
            .map_err(|e| ConvertError::Parse(format!("Failed to parse DOCX (docx-rs): {e}")))?;

        // Populate locale-specific footnote/endnote style IDs from docx styles
        notes.populate_style_ids(&docx.styles);

        let images = build_image_map(&docx);
        let hyperlinks = build_hyperlink_map(&docx);
        let numberings = build_numbering_map(&docx.numberings);
        let style_map = build_style_map(&docx.styles);
        let mut warnings: Vec<ConvertWarning> = Vec::new();

        let mut elements: Vec<TaggedElement> = Vec::new();
        let mut pages: Vec<Page> = Vec::new();
        let mut section_layout_index: usize = 0;
        for (idx, child) in docx.document.children.iter().enumerate() {
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| match child {
                docx_rs::DocumentChild::Paragraph(para) => {
                    let mut tagged = vec![convert_paragraph_element(
                        para,
                        &images,
                        &hyperlinks,
                        &style_map,
                        &notes,
                        &wraps,
                        &drawing_text_boxes,
                        &table_headers,
                        &vml_text_boxes,
                        &bidi,
                        &small_caps,
                    )];
                    // Inject math equations for this body child
                    let eqs = math.take(idx);
                    for eq in eqs {
                        tagged.push(TaggedElement::Plain(vec![Block::MathEquation(eq)]));
                    }
                    // Inject charts for this body child
                    let chs = chart_ctx.take(idx);
                    for ch in chs {
                        tagged.push(TaggedElement::Plain(vec![Block::Chart(ch)]));
                    }
                    tagged
                }
                docx_rs::DocumentChild::Table(table) => {
                    vec![TaggedElement::Plain(vec![Block::Table(convert_table(
                        table,
                        &images,
                        &hyperlinks,
                        &style_map,
                        &notes,
                        &wraps,
                        &drawing_text_boxes,
                        &table_headers,
                        &vml_text_boxes,
                        &bidi,
                        &small_caps,
                        0,
                    ))])]
                }
                docx_rs::DocumentChild::StructuredDataTag(sdt) => convert_sdt_children(
                    sdt,
                    &images,
                    &hyperlinks,
                    &style_map,
                    &notes,
                    &wraps,
                    &drawing_text_boxes,
                    &table_headers,
                    &vml_text_boxes,
                    &bidi,
                    &small_caps,
                ),
                _ => vec![TaggedElement::Plain(vec![])],
            }));

            match result {
                Ok(elems) => elements.extend(elems),
                Err(panic_info) => {
                    let detail = if let Some(s) = panic_info.downcast_ref::<String>() {
                        s.clone()
                    } else if let Some(s) = panic_info.downcast_ref::<&str>() {
                        (*s).to_string()
                    } else {
                        "unknown panic".to_string()
                    };
                    warnings.push(ConvertWarning::ParseSkipped {
                        format: "DOCX".to_string(),
                        reason: format!(
                            "upstream panic caught (docx-rs): element at index {idx}: {detail}"
                        ),
                    });
                }
            }

            if let docx_rs::DocumentChild::Paragraph(para) = child
                && let Some(section_prop) = para.property.section_property.as_ref()
            {
                let column_layout = match column_layouts.get(section_layout_index) {
                    Some(layout) => layout.clone(),
                    None => extract_column_layout_from_section_property(section_prop),
                };
                pages.push(Page::Flow(build_flow_page_from_section(
                    section_prop,
                    std::mem::take(&mut elements),
                    &numberings,
                    &header_footer_assets,
                    column_layout,
                    &mut warnings,
                )));
                section_layout_index += 1;
            }
        }

        let final_column_layout = match column_layouts.get(section_layout_index) {
            Some(layout) => layout.clone(),
            None => extract_column_layout_from_section_property(&docx.document.section_property),
        };
        pages.push(Page::Flow(build_flow_page_from_section(
            &docx.document.section_property,
            elements,
            &numberings,
            &header_footer_assets,
            final_column_layout,
            &mut warnings,
        )));

        Ok((
            Document {
                metadata,
                pages,
                styles: StyleSheet::default(),
            },
            warnings,
        ))
    }
}

fn convert_docx_header(header: &docx_rs::Header) -> Option<HeaderFooter> {
    let paragraphs = header
        .children
        .iter()
        .filter_map(|child| match child {
            docx_rs::HeaderChild::Paragraph(para) => Some(convert_hf_paragraph(para)),
            _ => None,
        })
        .collect::<Vec<_>>();
    if paragraphs.is_empty() {
        return None;
    }
    Some(HeaderFooter { paragraphs })
}

fn convert_docx_footer(footer: &docx_rs::Footer) -> Option<HeaderFooter> {
    let paragraphs = footer
        .children
        .iter()
        .filter_map(|child| match child {
            docx_rs::FooterChild::Paragraph(para) => Some(convert_hf_paragraph(para)),
            _ => None,
        })
        .collect::<Vec<_>>();
    if paragraphs.is_empty() {
        return None;
    }
    Some(HeaderFooter { paragraphs })
}

/// Extract the header for a section, preferring the default variant and falling back to
/// first/even variants when that is all the source document provides.
fn extract_docx_header(
    section_prop: &docx_rs::SectionProperty,
    assets: &HeaderFooterAssets,
) -> Option<HeaderFooter> {
    section_prop
        .header
        .as_ref()
        .and_then(|(_rid, header)| convert_docx_header(header))
        .or_else(|| {
            section_prop
                .header_reference
                .as_ref()
                .and_then(|reference| assets.headers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .first_header
                .as_ref()
                .and_then(|(_rid, header)| convert_docx_header(header))
        })
        .or_else(|| {
            section_prop
                .first_header_reference
                .as_ref()
                .and_then(|reference| assets.headers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .even_header
                .as_ref()
                .and_then(|(_rid, header)| convert_docx_header(header))
        })
        .or_else(|| {
            section_prop
                .even_header_reference
                .as_ref()
                .and_then(|reference| assets.headers.get(&reference.id).cloned())
        })
}

/// Extract the footer for a section, preferring the default variant and falling back to
/// first/even variants when that is all the source document provides.
fn extract_docx_footer(
    section_prop: &docx_rs::SectionProperty,
    assets: &HeaderFooterAssets,
) -> Option<HeaderFooter> {
    section_prop
        .footer
        .as_ref()
        .and_then(|(_rid, footer)| convert_docx_footer(footer))
        .or_else(|| {
            section_prop
                .footer_reference
                .as_ref()
                .and_then(|reference| assets.footers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .first_footer
                .as_ref()
                .and_then(|(_rid, footer)| convert_docx_footer(footer))
        })
        .or_else(|| {
            section_prop
                .first_footer_reference
                .as_ref()
                .and_then(|reference| assets.footers.get(&reference.id).cloned())
        })
        .or_else(|| {
            section_prop
                .even_footer
                .as_ref()
                .and_then(|(_rid, footer)| convert_docx_footer(footer))
        })
        .or_else(|| {
            section_prop
                .even_footer_reference
                .as_ref()
                .and_then(|reference| assets.footers.get(&reference.id).cloned())
        })
}

/// Convert a docx-rs Paragraph into a HeaderFooterParagraph.
/// Detects PAGE/NUMPAGES field codes within runs and emits page counter inlines.
fn convert_hf_paragraph(para: &docx_rs::Paragraph) -> HeaderFooterParagraph {
    let explicit_style = extract_paragraph_style(&para.property);
    let explicit_tab_overrides = extract_tab_stop_overrides(&para.property.tabs);
    let style = merge_paragraph_style(&explicit_style, explicit_tab_overrides.as_deref(), None);
    let mut elements: Vec<HFInline> = Vec::new();

    for child in &para.children {
        if let docx_rs::ParagraphChild::Run(run) = child {
            let run_style = extract_run_style(&run.run_property);
            extract_hf_run_elements(&run.children, &run_style, &mut elements);
        }
    }

    HeaderFooterParagraph { style, elements }
}

/// Extract inline elements from a run's children for header/footer use.
/// Recognizes text, tabs, and PAGE/NUMPAGES field codes.
fn extract_hf_run_elements(
    children: &[docx_rs::RunChild],
    style: &TextStyle,
    elements: &mut Vec<HFInline>,
) {
    let mut in_field = false;
    let mut field_inline: Option<HFInline> = None;
    let mut past_separate = false;

    for child in children {
        match child {
            docx_rs::RunChild::FieldChar(fc) => match fc.field_char_type {
                docx_rs::FieldCharType::Begin => {
                    in_field = true;
                    field_inline = None;
                    past_separate = false;
                }
                docx_rs::FieldCharType::Separate => {
                    past_separate = true;
                }
                docx_rs::FieldCharType::End => {
                    if let Some(inline) = field_inline.take() {
                        elements.push(inline);
                    }
                    in_field = false;
                    past_separate = false;
                }
                _ => {}
            },
            docx_rs::RunChild::InstrText(instr) => {
                if !in_field {
                    continue;
                }
                field_inline = match instr.as_ref() {
                    docx_rs::InstrText::PAGE(_) => Some(HFInline::PageNumber),
                    docx_rs::InstrText::NUMPAGES(_) => Some(HFInline::TotalPages),
                    _ => field_inline,
                };
            }
            docx_rs::RunChild::InstrTextString(s) => {
                if !in_field {
                    continue;
                }
                // After round-tripping through build/read_docx, InstrText::PAGE
                // becomes InstrTextString("PAGE"), and NUMPAGES likewise.
                let trimmed = s.trim();
                if trimmed.eq_ignore_ascii_case("page") {
                    field_inline = Some(HFInline::PageNumber);
                } else if trimmed.eq_ignore_ascii_case("numpages") {
                    field_inline = Some(HFInline::TotalPages);
                }
            }
            docx_rs::RunChild::Text(t) => {
                // Skip display values between separate and end
                if in_field && past_separate {
                    continue;
                }
                if !in_field && !t.text.is_empty() {
                    elements.push(HFInline::Run(Run {
                        text: t.text.clone(),
                        style: style.clone(),
                        href: None,
                        footnote: None,
                    }));
                }
            }
            docx_rs::RunChild::Tab(_) => {
                if !in_field {
                    elements.push(HFInline::Run(Run {
                        text: "\t".to_string(),
                        style: style.clone(),
                        href: None,
                        footnote: None,
                    }));
                }
            }
            _ => {}
        }
    }
}

/// Extract page size and margins from DOCX section properties.
fn extract_page_setup(section_prop: &docx_rs::SectionProperty) -> (PageSize, Margins) {
    let size = extract_page_size(&section_prop.page_size);
    let margins = extract_margins(&section_prop.page_margin);
    (size, margins)
}

/// Extract page size from docx-rs PageSize (which has private fields).
/// Uses serde serialization to access the private `w`, `h`, and `orient` fields.
/// Values in DOCX are in twips (1/20 of a point).
/// When orient is "landscape" and width < height, dimensions are swapped to ensure
/// landscape pages have width > height.
fn extract_page_size(page_size: &docx_rs::PageSize) -> PageSize {
    if let Ok(json) = serde_json::to_value(page_size) {
        let w = json.get("w").and_then(|v| v.as_f64()).unwrap_or(0.0);
        let h = json.get("h").and_then(|v| v.as_f64()).unwrap_or(0.0);
        let orient = json.get("orient").and_then(|v| v.as_str());
        if w > 0.0 && h > 0.0 {
            let mut width = w / 20.0; // twips to points
            let mut height = h / 20.0; // twips to points
            // If orient is landscape but dimensions are portrait-style, swap them
            if orient == Some("landscape") && width < height {
                std::mem::swap(&mut width, &mut height);
            }
            return PageSize { width, height };
        }
    }
    PageSize::default()
}

/// Extract margins from docx-rs PageMargin.
/// PageMargin fields are public i32 values in twips.
fn extract_margins(page_margin: &docx_rs::PageMargin) -> Margins {
    Margins {
        top: page_margin.top as f64 / 20.0,
        bottom: page_margin.bottom as f64 / 20.0,
        left: page_margin.left as f64 / 20.0,
        right: page_margin.right as f64 / 20.0,
    }
}

/// Extract content from a StructuredDataTag (SDT), processing its paragraph
/// and table children through the standard conversion pipeline.
/// SDTs are used for various structured content in DOCX, including Table of Contents.
#[allow(clippy::too_many_arguments)]
fn convert_sdt_children(
    sdt: &docx_rs::StructuredDataTag,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
) -> Vec<TaggedElement> {
    let mut result = Vec::new();
    for child in &sdt.children {
        match child {
            docx_rs::StructuredDataTagChild::Paragraph(para) => {
                result.push(convert_paragraph_element(
                    para,
                    images,
                    hyperlinks,
                    style_map,
                    notes,
                    wraps,
                    drawing_text_boxes,
                    table_headers,
                    vml_text_boxes,
                    bidi,
                    small_caps,
                ));
            }
            docx_rs::StructuredDataTagChild::Table(table) => {
                result.push(TaggedElement::Plain(vec![Block::Table(convert_table(
                    table,
                    images,
                    hyperlinks,
                    style_map,
                    notes,
                    wraps,
                    drawing_text_boxes,
                    table_headers,
                    vml_text_boxes,
                    bidi,
                    small_caps,
                    0,
                ))]));
            }
            docx_rs::StructuredDataTagChild::StructuredDataTag(nested) => {
                result.extend(convert_sdt_children(
                    nested,
                    images,
                    hyperlinks,
                    style_map,
                    notes,
                    wraps,
                    drawing_text_boxes,
                    table_headers,
                    vml_text_boxes,
                    bidi,
                    small_caps,
                ));
            }
            _ => {}
        }
    }
    result
}

/// Convert a docx-rs Paragraph into a TaggedElement.
/// If the paragraph has numbering, returns a `ListParagraph`; otherwise `Plain`.
#[allow(clippy::too_many_arguments)]
fn convert_paragraph_element(
    para: &docx_rs::Paragraph,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
) -> TaggedElement {
    let num_info = extract_num_info(para);

    // Build the paragraph IR
    let mut blocks = Vec::new();
    convert_paragraph_blocks(
        para,
        &mut blocks,
        images,
        hyperlinks,
        style_map,
        notes,
        wraps,
        drawing_text_boxes,
        table_headers,
        vml_text_boxes,
        bidi,
        small_caps,
    );

    match num_info {
        Some(info) => {
            // Extract the actual Paragraph from the blocks.
            // List paragraphs may also produce page breaks and images before the paragraph.
            let mut pre_blocks = Vec::new();
            let mut paragraph = None;
            for block in blocks {
                match block {
                    Block::Paragraph(p) if paragraph.is_none() => {
                        paragraph = Some(p);
                    }
                    _ => pre_blocks.push(block),
                }
            }
            if !pre_blocks.is_empty() {
                // If there were pre-blocks (page break, images), emit them as plain first.
                // We return the plain blocks — the caller will see them before the list paragraph.
                // For simplicity, we create a combined: Plain(pre) + ListParagraph.
                // But TaggedElement is a single value, so we need to handle this differently.
                // Actually, let's just emit them as plain first. The caller handles ordering.
                // Since we can only return one TaggedElement, fold the pre-blocks into the
                // paragraph by noting that list items in a list won't have page breaks.
                // For now, treat the paragraph as a plain block if it has pre-blocks.
                pre_blocks.push(Block::Paragraph(paragraph.unwrap_or_else(|| Paragraph {
                    style: ParagraphStyle::default(),
                    runs: Vec::new(),
                })));
                TaggedElement::Plain(pre_blocks)
            } else if let Some(p) = paragraph {
                TaggedElement::ListParagraph { info, paragraph: p }
            } else {
                TaggedElement::Plain(vec![])
            }
        }
        None => TaggedElement::Plain(blocks),
    }
}

/// Convert a docx-rs Paragraph to IR blocks, handling page breaks and inline images.
/// If the paragraph has `page_break_before`, a `Block::PageBreak` is emitted first.
/// Inline images within runs are extracted as separate `Block::Image` elements.
/// Style formatting from the document's style definitions is merged with explicit formatting.
#[allow(clippy::too_many_arguments)]
fn convert_paragraph_blocks(
    para: &docx_rs::Paragraph,
    out: &mut Vec<Block>,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
) {
    // Check bidi direction for this paragraph (must be called once per XML <w:p>)
    let is_rtl = bidi.next_is_bidi();

    // Emit page break before the paragraph if requested
    if para.property.page_break_before == Some(true) {
        out.push(Block::PageBreak);
    }

    // Look up the paragraph's referenced style
    let resolved_style = get_paragraph_style_id(&para.property)
        .and_then(|id| style_map.get(id))
        .or_else(|| style_map.get(DOC_DEFAULT_STYLE_ID));

    // Collect text runs and detect inline images
    let mut runs: Vec<Run> = Vec::new();
    let mut inline_images: Vec<Block> = Vec::new();
    let mut emitted_text_box_blocks: bool = false;

    for child in &para.children {
        match child {
            docx_rs::ParagraphChild::Run(run) => {
                // Advance smallCaps cursor for every <w:r> in body
                let is_small_caps: bool = small_caps.next_is_small_caps();

                // Check for footnote/endnote reference runs
                if is_note_reference_run(run, notes) {
                    if let Some(content) = notes.consume_next() {
                        runs.push(Run {
                            text: String::new(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: Some(content),
                        });
                    }
                    continue;
                }

                // Check for column breaks and embedded drawings in this run.
                let mut has_column_break = false;
                let mut text_box_blocks: Vec<Block> = Vec::new();
                for run_child in &run.children {
                    if let docx_rs::RunChild::Drawing(drawing) = run_child
                        && let Some(img_block) = extract_drawing_image(drawing, images, wraps)
                    {
                        inline_images.push(img_block);
                    }
                    if let docx_rs::RunChild::Drawing(drawing) = run_child {
                        text_box_blocks.extend(extract_drawing_text_box_blocks(
                            drawing,
                            images,
                            hyperlinks,
                            style_map,
                            notes,
                            wraps,
                            drawing_text_boxes,
                            table_headers,
                            vml_text_boxes,
                            bidi,
                            small_caps,
                        ));
                    }
                    if let docx_rs::RunChild::Shape(shape) = run_child {
                        let vml_text_box: VmlTextBoxInfo = vml_text_boxes.consume_next();
                        if let Some(floating_text_box) =
                            extract_vml_shape_text_box(shape, &vml_text_box)
                        {
                            text_box_blocks.push(Block::FloatingTextBox(floating_text_box));
                        } else {
                            text_box_blocks.extend(vml_text_box.into_blocks());
                        }

                        if let Some(img_block) = extract_shape_image(shape, images) {
                            inline_images.push(img_block);
                        }
                    }
                    if let docx_rs::RunChild::Break(br) = run_child
                        && is_column_break(br)
                    {
                        has_column_break = true;
                    }
                }

                if !text_box_blocks.is_empty() {
                    if !runs.is_empty() {
                        out.append(&mut inline_images);
                        push_paragraph_from_runs(out, para, resolved_style, is_rtl, &mut runs);
                    } else if !inline_images.is_empty() {
                        out.append(&mut inline_images);
                    }
                    emitted_text_box_blocks = true;
                    out.extend(text_box_blocks);
                }

                if has_column_break {
                    // Flush current runs as a paragraph before the column break
                    if !runs.is_empty() {
                        out.append(&mut inline_images);
                        push_paragraph_from_runs(out, para, resolved_style, is_rtl, &mut runs);
                    }
                    out.push(Block::ColumnBreak);

                    // Still extract any text from this run (after the break)
                    let text = extract_run_text_skip_column_breaks(run);
                    if !text.is_empty() {
                        let mut explicit_style: TextStyle = extract_run_style(&run.run_property);
                        if is_small_caps {
                            explicit_style.small_caps = Some(true);
                        }
                        runs.push(Run {
                            text,
                            style: merge_text_style(&explicit_style, resolved_style),
                            href: None,
                            footnote: None,
                        });
                    }
                } else {
                    // Extract text from the run
                    let text = extract_run_text(run);
                    if !text.is_empty() {
                        let mut explicit_style: TextStyle = extract_run_style(&run.run_property);
                        if is_small_caps {
                            explicit_style.small_caps = Some(true);
                        }
                        runs.push(Run {
                            text,
                            style: merge_text_style(&explicit_style, resolved_style),
                            href: None,
                            footnote: None,
                        });
                    }
                }
            }
            docx_rs::ParagraphChild::Hyperlink(hyperlink) => {
                // Resolve the hyperlink URL from document relationships
                let href = resolve_hyperlink_url(hyperlink, hyperlinks);

                // Extract runs from inside the hyperlink element
                for hchild in &hyperlink.children {
                    if let docx_rs::ParagraphChild::Run(run) = hchild {
                        let hl_small_caps: bool = small_caps.next_is_small_caps();
                        let text = extract_run_text(run);
                        if !text.is_empty() {
                            let mut explicit_style: TextStyle =
                                extract_run_style(&run.run_property);
                            if hl_small_caps {
                                explicit_style.small_caps = Some(true);
                            }
                            runs.push(Run {
                                text,
                                style: merge_text_style(&explicit_style, resolved_style),
                                href: href.clone(),
                                footnote: None,
                            });
                        }
                    }
                }
            }
            _ => {}
        }
    }

    // Emit image blocks before the paragraph (inline images are block-level in our IR)
    out.extend(inline_images);

    if !runs.is_empty() || !emitted_text_box_blocks {
        push_paragraph_from_runs(out, para, resolved_style, is_rtl, &mut runs);
    }
}

fn push_paragraph_from_runs(
    out: &mut Vec<Block>,
    para: &docx_rs::Paragraph,
    resolved_style: Option<&ResolvedStyle>,
    is_rtl: bool,
    runs: &mut Vec<Run>,
) {
    let explicit_para_style = extract_paragraph_style(&para.property);
    let explicit_tab_overrides = extract_tab_stop_overrides(&para.property.tabs);
    let mut style = merge_paragraph_style(
        &explicit_para_style,
        explicit_tab_overrides.as_deref(),
        resolved_style,
    );
    if is_rtl {
        style.direction = Some(TextDirection::Rtl);
    }
    out.push(Block::Paragraph(Paragraph {
        style,
        runs: std::mem::take(runs),
    }));
}

/// Convert EMU (English Metric Units) to points for signed values (position offsets).
fn emu_to_pt_signed(emu: i32) -> f64 {
    emu as f64 / 12700.0
}

/// Extract an image from a Drawing element if it contains a Pic with matching image data.
/// Anchor images (floating) are returned as `Block::FloatingImage` with wrap mode from context.
/// Inline images are returned as `Block::Image`.
fn extract_drawing_image(
    drawing: &docx_rs::Drawing,
    images: &ImageMap,
    wraps: &WrapContext,
) -> Option<Block> {
    let pic = match &drawing.data {
        Some(docx_rs::DrawingData::Pic(pic)) => pic,
        _ => return None,
    };

    // Look up image data by relationship ID
    let data = images.get(&pic.id)?;

    let (w_emu, h_emu) = pic.size;
    let width = if w_emu > 0 {
        Some(emu_to_pt(w_emu))
    } else {
        None
    };
    let height = if h_emu > 0 {
        Some(emu_to_pt(h_emu))
    } else {
        None
    };

    let image_data = ImageData {
        data: data.clone(),
        format: ImageFormat::Png, // docx-rs converts all images to PNG
        width,
        height,
        crop: None,
    };

    // Check if this is an anchor (floating) image
    if pic.position_type == docx_rs::DrawingPositionType::Anchor {
        let wrap_mode = wraps.consume_next();

        // Extract position offsets from the Pic
        let offset_x = match pic.position_h {
            docx_rs::DrawingPosition::Offset(emu) => emu_to_pt_signed(emu),
            docx_rs::DrawingPosition::Align(_) => 0.0,
        };
        let offset_y = match pic.position_v {
            docx_rs::DrawingPosition::Offset(emu) => emu_to_pt_signed(emu),
            docx_rs::DrawingPosition::Align(_) => 0.0,
        };

        Some(Block::FloatingImage(FloatingImage {
            image: image_data,
            wrap_mode,
            offset_x,
            offset_y,
        }))
    } else {
        Some(Block::Image(image_data))
    }
}

fn extract_shape_image(shape: &docx_rs::Shape, images: &ImageMap) -> Option<Block> {
    let image_id = shape.image_data.as_ref()?.id.as_str();
    let data = images.get(image_id)?;

    let width = extract_vml_style_dimension(shape.style.as_deref(), "width");
    let height = extract_vml_style_dimension(shape.style.as_deref(), "height");

    Some(Block::Image(ImageData {
        data: data.clone(),
        format: ImageFormat::Png,
        width,
        height,
        crop: None,
    }))
}

fn extract_vml_shape_text_box(
    shape: &docx_rs::Shape,
    text_box: &VmlTextBoxInfo,
) -> Option<FloatingTextBox> {
    if text_box.paragraphs.is_empty() {
        return None;
    }

    let style = shape.style.as_deref()?;
    if !is_positioned_vml_text_box(style) {
        return None;
    }

    let width = extract_vml_style_length(Some(style), "width")?;
    let height = extract_vml_style_length(Some(style), "height")?;
    let offset_x = extract_vml_style_length(Some(style), "margin-left")
        .or_else(|| extract_vml_style_length(Some(style), "left"))
        .unwrap_or(0.0);
    let offset_y = extract_vml_style_length(Some(style), "margin-top")
        .or_else(|| extract_vml_style_length(Some(style), "top"))
        .unwrap_or(0.0);
    let wrap_mode = text_box
        .wrap_mode
        .or_else(|| extract_vml_style_wrap_mode(Some(style)))
        .unwrap_or(WrapMode::Square);

    Some(FloatingTextBox {
        content: text_box.clone().into_blocks(),
        wrap_mode,
        width,
        height,
        offset_x,
        offset_y,
    })
}

fn is_positioned_vml_text_box(style: &str) -> bool {
    has_vml_style_value(style, "position", "absolute")
        || extract_vml_style_length(Some(style), "margin-left").is_some()
        || extract_vml_style_length(Some(style), "margin-top").is_some()
}

fn has_vml_style_value(style: &str, key: &str, expected: &str) -> bool {
    extract_vml_style_value(style, key)
        .map(|value| value.eq_ignore_ascii_case(expected))
        .unwrap_or(false)
}

fn extract_vml_style_wrap_mode(style: Option<&str>) -> Option<WrapMode> {
    let value = extract_vml_style_value(style?, "mso-wrap-style")?;
    match value.to_ascii_lowercase().as_str() {
        "square" => Some(WrapMode::Square),
        "none" => Some(WrapMode::None),
        "tight" | "through" => Some(WrapMode::Tight),
        "top-and-bottom" | "topandbottom" => Some(WrapMode::TopAndBottom),
        _ => None,
    }
}

fn extract_vml_style_value(style: &str, key: &str) -> Option<String> {
    for part in style.split(';') {
        let Some((name, value)) = part.split_once(':') else {
            continue;
        };
        if name.trim() == key {
            return Some(value.trim().to_string());
        }
    }

    None
}

fn extract_vml_style_length(style: Option<&str>, key: &str) -> Option<f64> {
    let value = extract_vml_style_value(style?, key)?;
    let value = value.trim();
    if let Some(raw) = value.strip_suffix("pt") {
        return raw.trim().parse::<f64>().ok();
    }
    if let Some(raw) = value.strip_suffix("px") {
        return raw.trim().parse::<f64>().ok().map(|px| px * 72.0 / 96.0);
    }

    None
}

fn extract_vml_style_dimension(style: Option<&str>, key: &str) -> Option<f64> {
    let style = style?;
    for part in style.split(';') {
        let Some((name, value)) = part.split_once(':') else {
            continue;
        };
        if name.trim() != key {
            continue;
        }

        let value = value.trim();
        if let Some(raw) = value.strip_suffix("pt") {
            return raw.trim().parse::<f64>().ok();
        }
        if let Some(raw) = value.strip_suffix("px") {
            return raw.trim().parse::<f64>().ok().map(|px| px * 72.0 / 96.0);
        }
        if let Ok(points) = value.parse::<f64>() {
            return Some(points);
        }
    }

    None
}

#[allow(clippy::too_many_arguments)]
fn extract_drawing_text_box_blocks(
    drawing: &docx_rs::Drawing,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
) -> Vec<Block> {
    let Some(docx_rs::DrawingData::TextBox(text_box)) = &drawing.data else {
        return Vec::new();
    };

    let layout: DrawingTextBoxInfo = drawing_text_boxes.consume_next();
    let mut blocks: Vec<Block> = Vec::new();
    for child in &text_box.children {
        match child {
            docx_rs::TextBoxContentChild::Paragraph(para) => convert_paragraph_blocks(
                para,
                &mut blocks,
                images,
                hyperlinks,
                style_map,
                notes,
                wraps,
                drawing_text_boxes,
                table_headers,
                vml_text_boxes,
                bidi,
                small_caps,
            ),
            docx_rs::TextBoxContentChild::Table(table) => {
                blocks.push(Block::Table(convert_table(
                    table,
                    images,
                    hyperlinks,
                    style_map,
                    notes,
                    wraps,
                    drawing_text_boxes,
                    table_headers,
                    vml_text_boxes,
                    bidi,
                    small_caps,
                    0,
                )));
            }
        }
    }

    if text_box.position_type == docx_rs::DrawingPositionType::Anchor {
        let wrap_mode = wraps.consume_next();
        let offset_x = match text_box.position_h {
            docx_rs::DrawingPosition::Offset(emu) => emu_to_pt_signed(emu),
            docx_rs::DrawingPosition::Align(_) => 0.0,
        };
        let offset_y = match text_box.position_v {
            docx_rs::DrawingPosition::Offset(emu) => emu_to_pt_signed(emu),
            docx_rs::DrawingPosition::Align(_) => 0.0,
        };
        let (width, height) = resolve_drawing_text_box_size(text_box, layout);

        vec![Block::FloatingTextBox(FloatingTextBox {
            content: blocks,
            wrap_mode,
            width,
            height,
            offset_x,
            offset_y,
        })]
    } else {
        blocks
    }
}

fn resolve_drawing_text_box_size(
    text_box: &docx_rs::TextBox,
    layout: DrawingTextBoxInfo,
) -> (f64, f64) {
    let width = layout.width_pt.unwrap_or_else(|| {
        if text_box.size.0 > 0 {
            emu_to_pt(text_box.size.0)
        } else {
            0.0
        }
    });
    let height = layout.height_pt.unwrap_or_else(|| {
        if text_box.size.1 > 0 {
            emu_to_pt(text_box.size.1)
        } else {
            0.0
        }
    });

    (width, height)
}

/// Extract paragraph-level formatting from docx-rs ParagraphProperty.
fn extract_paragraph_style(prop: &docx_rs::ParagraphProperty) -> ParagraphStyle {
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
        space_before,
        space_after,
        heading_level: None,
        direction: None, // Set by BidiContext after style merge
        tab_stops,
    }
}

/// Extract indentation values from docx-rs Indent.
/// All values in docx-rs are in twips; convert to points (÷20).
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

/// Extract line spacing from docx-rs LineSpacing (private fields, use serde).
///
/// OOXML line spacing:
/// - Auto: `line` is in 240ths of a line (240=single, 360=1.5×, 480=double)
/// - Exact/AtLeast: `line` is in twips (÷20 → points)
/// - `before`/`after`: twips (÷20 → points)
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
            _ => {
                // Auto: 240ths of a line
                Some(LineSpacing::Proportional(line / 240.0))
            }
        }
    });

    (line_spacing, space_before, space_after)
}

/// Extract tab stops from paragraph properties.
/// docx-rs Tab has `pos` in twips and `val`/`leader` as enums.
fn extract_tab_stops(tabs: &[docx_rs::Tab]) -> Option<Vec<TabStop>> {
    let tab_overrides = extract_tab_stop_overrides(tabs)?;
    let mut tab_stops: Vec<TabStop> = Vec::new();
    apply_tab_stop_overrides(&mut tab_stops, &tab_overrides);
    Some(tab_stops)
}

fn extract_tab_stop_overrides(tabs: &[docx_rs::Tab]) -> Option<Vec<TabStopOverride>> {
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

fn extract_margin_side_points(side_json: &serde_json::Value) -> Option<f64> {
    let width_type = side_json
        .get("widthType")
        .and_then(|v| v.as_str())
        .unwrap_or("dxa");
    let value = side_json.get("val").and_then(|v| v.as_f64())?;

    match width_type {
        "dxa" => Some(value / 20.0),
        _ => None,
    }
}

fn extract_insets_from_margins_json(margins_json: &serde_json::Value) -> Option<Insets> {
    let top = margins_json.get("top").and_then(extract_margin_side_points);
    let right = margins_json
        .get("right")
        .and_then(extract_margin_side_points);
    let bottom = margins_json
        .get("bottom")
        .and_then(extract_margin_side_points);
    let left = margins_json
        .get("left")
        .and_then(extract_margin_side_points);

    if top.is_none() && right.is_none() && bottom.is_none() && left.is_none() {
        return None;
    }

    Some(Insets {
        top: top.unwrap_or_default(),
        right: right.unwrap_or_default(),
        bottom: bottom.unwrap_or_default(),
        left: left.unwrap_or_default(),
    })
}

fn extract_table_alignment(prop_json: Option<&serde_json::Value>) -> Option<Alignment> {
    prop_json
        .and_then(|j| j.get("justification"))
        .and_then(|v| v.as_str())
        .and_then(|value| match value {
            "center" => Some(Alignment::Center),
            "right" | "end" => Some(Alignment::Right),
            _ => None,
        })
}

fn extract_table_default_cell_padding(prop_json: Option<&serde_json::Value>) -> Option<Insets> {
    prop_json
        .and_then(|j| j.get("margins"))
        .and_then(extract_insets_from_margins_json)
}

fn extract_cell_padding(
    prop_json: Option<&serde_json::Value>,
    inherited_padding: Option<Insets>,
) -> Option<Insets> {
    let margins_json = prop_json.and_then(|j| j.get("margins"))?;
    extract_insets_from_margins_json(margins_json)?;
    let mut merged_padding = inherited_padding.unwrap_or_default();

    if let Some(top) = margins_json.get("top").and_then(extract_margin_side_points) {
        merged_padding.top = top;
    }
    if let Some(right) = margins_json
        .get("right")
        .and_then(extract_margin_side_points)
    {
        merged_padding.right = right;
    }
    if let Some(bottom) = margins_json
        .get("bottom")
        .and_then(extract_margin_side_points)
    {
        merged_padding.bottom = bottom;
    }
    if let Some(left) = margins_json
        .get("left")
        .and_then(extract_margin_side_points)
    {
        merged_padding.left = left;
    }

    Some(merged_padding)
}

fn extract_table_cell_width(prop_json: Option<&serde_json::Value>) -> Option<f64> {
    let width_json = prop_json.and_then(|j| j.get("width"))?;
    let width_type = width_json
        .get("widthType")
        .and_then(|v| v.as_str())
        .unwrap_or("dxa");
    let width = width_json.get("width").and_then(|v| v.as_f64())?;

    match width_type {
        "dxa" => Some(width / 20.0),
        _ => None,
    }
}

#[allow(clippy::too_many_arguments)]
/// Convert a docx-rs Table to an IR Table.
///
/// Handles:
/// - Column widths from the table grid (twips → points)
/// - Cell content (paragraphs with formatted text)
/// - Horizontal merging via gridSpan (colspan)
/// - Vertical merging via vMerge restart/continue (rowspan)
/// - Cell background color via shading
/// - Cell borders
fn convert_table(
    table: &docx_rs::Table,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
    depth: usize,
) -> Table {
    let header_info = table_headers.consume_next();
    let table_prop_json = serde_json::to_value(&table.property).ok();
    let alignment = extract_table_alignment(table_prop_json.as_ref());
    let default_cell_padding = extract_table_default_cell_padding(table_prop_json.as_ref());

    // First pass: extract raw rows with vmerge info for rowspan calculation
    let raw_rows = extract_raw_rows(
        table,
        images,
        hyperlinks,
        style_map,
        notes,
        wraps,
        drawing_text_boxes,
        table_headers,
        vml_text_boxes,
        bidi,
        small_caps,
        depth,
        default_cell_padding,
    );

    let column_widths: Vec<f64> = if table.grid.is_empty() {
        derive_column_widths_from_cells(&raw_rows).unwrap_or_default()
    } else {
        table.grid.iter().map(|&w| w as f64 / 20.0).collect()
    };

    // Second pass: resolve vertical merges into rowspan values and build IR rows
    let rows = resolve_vmerge_and_build_rows(&raw_rows);

    Table {
        rows,
        column_widths,
        header_row_count: header_info.repeat_rows.min(table.rows.len()),
        alignment,
        default_cell_padding,
        use_content_driven_row_heights: false,
    }
}

/// Intermediate cell representation for vmerge resolution.
struct RawCell {
    content: Vec<Block>,
    col_span: u32,
    col_index: usize,
    preferred_width: Option<f64>,
    vmerge: Option<String>, // "restart", "continue", or None
    border: Option<CellBorder>,
    background: Option<Color>,
    vertical_align: Option<CellVerticalAlign>,
    padding: Option<Insets>,
}

struct RawRow {
    cells: Vec<RawCell>,
    height: Option<f64>,
}

#[allow(clippy::too_many_arguments)]
/// Extract raw rows from a docx-rs Table, tracking column indices and vmerge state.
fn extract_raw_rows(
    table: &docx_rs::Table,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
    depth: usize,
    default_cell_padding: Option<Insets>,
) -> Vec<RawRow> {
    let mut raw_rows = Vec::new();

    for table_child in &table.rows {
        let docx_rs::TableChild::TableRow(row) = table_child;
        let row_prop_json = serde_json::to_value(&row.property).ok();
        // Typst row tracks are exact sizes, so only preserve DOCX row heights
        // when Word marks them as exact. AtLeast would require min-content sizing.
        let height = row_prop_json
            .as_ref()
            .filter(|j| j.get("heightRule").and_then(|v| v.as_str()) == Some("exact"))
            .and_then(|j| j.get("rowHeight"))
            .and_then(|v| v.as_f64());
        let mut cells = Vec::new();
        let mut col_index: usize = 0;

        for row_child in &row.cells {
            let docx_rs::TableRowChild::TableCell(cell) = row_child;

            let prop_json = serde_json::to_value(&cell.property).ok();
            let grid_span = prop_json
                .as_ref()
                .and_then(|j| j.get("gridSpan"))
                .and_then(|v| v.as_u64())
                .unwrap_or(1) as u32;

            let vmerge = prop_json
                .as_ref()
                .and_then(|j| j.get("verticalMerge"))
                .and_then(|v| v.as_str())
                .map(String::from);
            let preferred_width = extract_table_cell_width(prop_json.as_ref());

            let content = extract_cell_content(
                cell,
                images,
                hyperlinks,
                style_map,
                notes,
                wraps,
                drawing_text_boxes,
                table_headers,
                vml_text_boxes,
                bidi,
                small_caps,
                depth,
            );
            let border = prop_json
                .as_ref()
                .and_then(|j| j.get("borders"))
                .and_then(extract_cell_borders);
            let background = prop_json
                .as_ref()
                .and_then(|j| j.get("shading"))
                .and_then(extract_cell_shading);

            let vertical_align: Option<CellVerticalAlign> = prop_json
                .as_ref()
                .and_then(|j| j.get("verticalAlign"))
                .and_then(|v| v.as_str())
                .and_then(|s| match s {
                    "center" => Some(CellVerticalAlign::Center),
                    "bottom" => Some(CellVerticalAlign::Bottom),
                    _ => None, // "top" is default, skip
                });
            let padding = extract_cell_padding(prop_json.as_ref(), default_cell_padding);

            cells.push(RawCell {
                content,
                col_span: grid_span,
                col_index,
                preferred_width,
                vmerge,
                border,
                background,
                vertical_align,
                padding,
            });

            col_index += grid_span as usize;
        }

        raw_rows.push(RawRow { cells, height });
    }

    raw_rows
}

fn derive_column_widths_from_cells(raw_rows: &[RawRow]) -> Option<Vec<f64>> {
    let num_cols = raw_rows
        .iter()
        .flat_map(|row| {
            row.cells
                .iter()
                .map(|cell| cell.col_index + cell.col_span as usize)
        })
        .max()
        .unwrap_or(0);

    if num_cols == 0 {
        return None;
    }

    let mut widths: Vec<f64> = vec![0.0; num_cols];
    let mut saw_width = false;

    for row in raw_rows {
        for cell in &row.cells {
            let Some(preferred_width) = cell.preferred_width else {
                continue;
            };
            if cell.col_span == 0 {
                continue;
            }

            let per_column_width = preferred_width / cell.col_span as f64;
            for width in widths
                .iter_mut()
                .skip(cell.col_index)
                .take(cell.col_span as usize)
            {
                *width = width.max(per_column_width);
            }
            saw_width = true;
        }
    }

    saw_width.then_some(widths)
}

/// Resolve vertical merges: compute rowspan for "restart" cells and skip "continue" cells.
fn resolve_vmerge_and_build_rows(raw_rows: &[RawRow]) -> Vec<TableRow> {
    let mut rows = Vec::new();

    for (row_idx, raw_row) in raw_rows.iter().enumerate() {
        let mut cells = Vec::new();

        for raw_cell in &raw_row.cells {
            match raw_cell.vmerge.as_deref() {
                Some("continue") => {
                    // Skip continue cells — they are part of a vertical merge above
                    continue;
                }
                Some("restart") => {
                    // Count how many consecutive "continue" cells follow in the same column
                    let row_span = count_vmerge_span(raw_rows, row_idx, raw_cell.col_index);
                    cells.push(TableCell {
                        content: raw_cell.content.clone(),
                        col_span: raw_cell.col_span,
                        row_span,
                        border: raw_cell.border.clone(),
                        background: raw_cell.background,
                        data_bar: None,
                        icon_text: None,
                        vertical_align: raw_cell.vertical_align,
                        padding: raw_cell.padding,
                    });
                }
                _ => {
                    // Normal cell (no vmerge)
                    cells.push(TableCell {
                        content: raw_cell.content.clone(),
                        col_span: raw_cell.col_span,
                        row_span: 1,
                        border: raw_cell.border.clone(),
                        background: raw_cell.background,
                        data_bar: None,
                        icon_text: None,
                        vertical_align: raw_cell.vertical_align,
                        padding: raw_cell.padding,
                    });
                }
            }
        }

        rows.push(TableRow {
            cells,
            height: raw_row.height,
        });
    }

    rows
}

/// Count the vertical merge span starting from a "restart" cell.
/// Looks at rows below `start_row` for "continue" cells at the same column index.
fn count_vmerge_span(raw_rows: &[RawRow], start_row: usize, col_index: usize) -> u32 {
    let mut span = 1u32;
    for row in raw_rows.iter().skip(start_row + 1) {
        let has_continue = row
            .cells
            .iter()
            .any(|c| c.col_index == col_index && c.vmerge.as_deref() == Some("continue"));
        if has_continue {
            span += 1;
        } else {
            break;
        }
    }
    span
}

#[allow(clippy::too_many_arguments)]
/// Extract cell content (paragraphs) from a docx-rs TableCell.
fn extract_cell_content(
    cell: &docx_rs::TableCell,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
    drawing_text_boxes: &DrawingTextBoxContext,
    table_headers: &TableHeaderContext,
    vml_text_boxes: &VmlTextBoxContext,
    bidi: &BidiContext,
    small_caps: &SmallCapsContext,
    depth: usize,
) -> Vec<Block> {
    let mut blocks = Vec::new();
    for content in &cell.children {
        match content {
            docx_rs::TableCellContent::Paragraph(para) => {
                convert_paragraph_blocks(
                    para,
                    &mut blocks,
                    images,
                    hyperlinks,
                    style_map,
                    notes,
                    wraps,
                    drawing_text_boxes,
                    table_headers,
                    vml_text_boxes,
                    bidi,
                    small_caps,
                );
            }
            docx_rs::TableCellContent::Table(nested_table) => {
                if depth < MAX_TABLE_DEPTH {
                    blocks.push(Block::Table(convert_table(
                        nested_table,
                        images,
                        hyperlinks,
                        style_map,
                        notes,
                        wraps,
                        drawing_text_boxes,
                        table_headers,
                        vml_text_boxes,
                        bidi,
                        small_caps,
                        depth + 1,
                    )));
                }
                // Silently skip nested tables beyond MAX_TABLE_DEPTH
            }
            _ => {}
        }
    }
    blocks
}

/// Extract cell borders from the serialized "borders" JSON object.
/// Border size in docx-rs is in eighths of a point; convert to points (÷8).
fn extract_cell_borders(borders_json: &serde_json::Value) -> Option<CellBorder> {
    if borders_json.is_null() {
        return None;
    }

    let extract_side = |key: &str| -> Option<BorderSide> {
        let side = borders_json.get(key)?;
        if side.is_null() {
            return None;
        }
        let border_type = side
            .get("borderType")
            .and_then(|v| v.as_str())
            .unwrap_or("none");
        if border_type == "none" || border_type == "nil" {
            return None;
        }
        let size = side.get("size").and_then(|v| v.as_f64()).unwrap_or(0.0);
        let color_hex = side
            .get("color")
            .and_then(|v| v.as_str())
            .unwrap_or("000000");
        let color = parse_hex_color(color_hex).unwrap_or(Color::black());
        let style = match border_type {
            "dashed" | "dashSmallGap" => BorderLineStyle::Dashed,
            "dotted" => BorderLineStyle::Dotted,
            "dashDotStroked" | "dotDash" => BorderLineStyle::DashDot,
            "dotDotDash" => BorderLineStyle::DashDotDot,
            "double"
            | "thinThickSmallGap"
            | "thickThinSmallGap"
            | "thinThickMediumGap"
            | "thickThinMediumGap"
            | "thinThickLargeGap"
            | "thickThinLargeGap"
            | "thinThickThinSmallGap"
            | "thinThickThinMediumGap"
            | "thinThickThinLargeGap"
            | "triple" => BorderLineStyle::Double,
            _ => BorderLineStyle::Solid,
        };
        Some(BorderSide {
            width: size / 8.0, // eighths of a point → points
            color,
            style,
        })
    };

    let top = extract_side("top");
    let bottom = extract_side("bottom");
    let left = extract_side("left");
    let right = extract_side("right");

    if top.is_none() && bottom.is_none() && left.is_none() && right.is_none() {
        return None;
    }

    Some(CellBorder {
        top,
        bottom,
        left,
        right,
    })
}

/// Extract background color from the serialized "shading" JSON object.
fn extract_cell_shading(shading_json: &serde_json::Value) -> Option<Color> {
    if shading_json.is_null() {
        return None;
    }
    let fill = shading_json.get("fill").and_then(|v| v.as_str())?;
    // Skip auto/transparent fills
    if fill == "auto" || fill == "FFFFFF" || fill.is_empty() {
        return None;
    }
    parse_hex_color(fill)
}

/// Extract inline text style from a docx-rs RunProperty.
///
/// docx-rs types serialize directly as their inner value (e.g. Bold → `true`,
/// Sz → `24`, Color → `"FF0000"`), so JSON extraction works for both explicit
/// run properties and docDefaults run properties.
fn extract_run_style(rp: &docx_rs::RunProperty) -> TextStyle {
    let json = serde_json::to_value(rp).unwrap_or(serde_json::Value::Null);
    extract_run_style_from_json(&json)
}

/// Extract inline text style from a serialized RunProperty-like JSON object.
fn extract_run_style_from_json(rp: &serde_json::Value) -> TextStyle {
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
            .map(|half_points| half_points / 2.0),
        color: rp
            .get("color")
            .and_then(serde_json::Value::as_str)
            .and_then(parse_hex_color),
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
        // smallCaps is not exposed by docx-rs; set via SmallCapsContext XML scan
        small_caps: None,
        // character_spacing is in twips (1/20 pt); convert to points
        letter_spacing: rp
            .get("characterSpacing")
            .and_then(serde_json::Value::as_i64)
            .map(|twips| twips as f64 / 20.0),
    }
}

fn json_bool_or_val(value: &serde_json::Value) -> Option<bool> {
    value
        .as_bool()
        .or_else(|| value.get("val").and_then(serde_json::Value::as_bool))
}

/// Extract document-level default text style from styles.xml docDefaults.
fn extract_doc_default_text_style(styles: &docx_rs::Styles) -> TextStyle {
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

/// Map OOXML named highlight colors to RGB values.
/// The 16 named colors are defined in the ECMA-376 spec (ST_HighlightColor).
fn resolve_highlight_color(name: &str) -> Option<Color> {
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
        _ => None, // "none" or unrecognized
    }
}

/// Parse a 6-character hex color string (e.g. "FF0000") to an IR Color.
fn parse_hex_color(hex: &str) -> Option<Color> {
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some(Color::new(r, g, b))
}

/// Resolve a hyperlink's URL from document relationships.
/// For external hyperlinks, looks up the relationship ID in the hyperlink map.
/// For anchor hyperlinks (internal bookmarks), returns None.
fn resolve_hyperlink_url(
    hyperlink: &docx_rs::Hyperlink,
    hyperlinks: &HyperlinkMap,
) -> Option<String> {
    match &hyperlink.link {
        docx_rs::HyperlinkData::External { rid, path } => {
            // First try the path (populated during writing),
            // then fall back to the relationship map (populated during reading)
            if !path.is_empty() {
                Some(path.clone())
            } else {
                hyperlinks.get(rid).cloned()
            }
        }
        docx_rs::HyperlinkData::Anchor { .. } => None, // internal bookmark, skip
    }
}

/// Check if a docx-rs Break is a column break.
/// Break.break_type is private, so we use serde to extract the value.
fn is_column_break(br: &docx_rs::Break) -> bool {
    serde_json::to_value(br)
        .ok()
        .and_then(|v| {
            v.get("breakType")
                .and_then(|bt| bt.as_str().map(|s| s == "column"))
        })
        .unwrap_or(false)
}

/// Extract text content from a docx-rs Run, skipping column breaks.
/// Column breaks are handled separately as Block::ColumnBreak.
fn extract_run_text_skip_column_breaks(run: &docx_rs::Run) -> String {
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

/// Extract text content from a docx-rs Run.
fn extract_run_text(run: &docx_rs::Run) -> String {
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ir::*;
    use std::collections::BTreeMap;
    use std::io::Cursor;

    /// Helper: build a minimal DOCX as bytes using docx-rs builder.
    fn build_docx_bytes(paragraphs: Vec<docx_rs::Paragraph>) -> Vec<u8> {
        let mut docx = docx_rs::Docx::new();
        for p in paragraphs {
            docx = docx.add_paragraph(p);
        }
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: build a DOCX with custom page size and margins.
    fn build_docx_bytes_with_page_setup(
        paragraphs: Vec<docx_rs::Paragraph>,
        width_twips: u32,
        height_twips: u32,
        margin_top: i32,
        margin_bottom: i32,
        margin_left: i32,
        margin_right: i32,
    ) -> Vec<u8> {
        let mut docx = docx_rs::Docx::new()
            .page_size(width_twips, height_twips)
            .page_margin(
                docx_rs::PageMargin::new()
                    .top(margin_top)
                    .bottom(margin_bottom)
                    .left(margin_left)
                    .right(margin_right),
            );
        for p in paragraphs {
            docx = docx.add_paragraph(p);
        }
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    // ----- Basic parsing tests -----

    #[test]
    fn test_parse_empty_docx() {
        let data = build_docx_bytes(vec![]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        // An empty DOCX should produce a document with one FlowPage and no content blocks
        assert_eq!(doc.pages.len(), 1);
        match &doc.pages[0] {
            Page::Flow(page) => {
                assert!(page.content.is_empty());
            }
            _ => panic!("Expected FlowPage"),
        }
    }

    #[test]
    fn test_parse_single_paragraph() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello, world!")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(doc.pages.len(), 1);
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        assert_eq!(page.content.len(), 1);
        match &page.content[0] {
            Block::Paragraph(para) => {
                assert_eq!(para.runs.len(), 1);
                assert_eq!(para.runs[0].text, "Hello, world!");
            }
            _ => panic!("Expected Paragraph block"),
        }
    }

    #[test]
    fn test_parse_multiple_paragraphs() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("First paragraph")),
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Second paragraph")),
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Third paragraph")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        assert_eq!(page.content.len(), 3);

        let texts: Vec<&str> = page
            .content
            .iter()
            .map(|b| match b {
                Block::Paragraph(p) => p.runs[0].text.as_str(),
                _ => panic!("Expected Paragraph"),
            })
            .collect();
        assert_eq!(
            texts,
            vec!["First paragraph", "Second paragraph", "Third paragraph"]
        );
    }

    #[test]
    fn test_parse_paragraph_with_multiple_runs() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Hello, "))
                .add_run(docx_rs::Run::new().add_text("beautiful "))
                .add_run(docx_rs::Run::new().add_text("world!")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs.len(), 3);
        assert_eq!(para.runs[0].text, "Hello, ");
        assert_eq!(para.runs[1].text, "beautiful ");
        assert_eq!(para.runs[2].text, "world!");
    }

    #[test]
    fn test_parse_empty_paragraph() {
        let data = build_docx_bytes(vec![docx_rs::Paragraph::new()]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        // An empty paragraph should still be present (it may have no runs)
        assert_eq!(page.content.len(), 1);
        match &page.content[0] {
            Block::Paragraph(para) => {
                assert!(para.runs.is_empty());
            }
            _ => panic!("Expected Paragraph block"),
        }
    }

    // ----- Page setup tests -----

    #[test]
    fn test_default_page_size_is_used() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        // docx-rs default: 11906 x 16838 twips (A4)
        // = 595.3 x 841.9 pt
        assert!(page.size.width > 0.0);
        assert!(page.size.height > 0.0);
    }

    #[test]
    fn test_custom_page_size_extracted() {
        // A5 page: 148mm x 210mm
        // In twips: 8392 x 11907 (1 pt = 20 twips)
        let width_twips: u32 = 8392;
        let height_twips: u32 = 11907;
        let data = build_docx_bytes_with_page_setup(
            vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test"))],
            width_twips,
            height_twips,
            1440,
            1440,
            1440,
            1440,
        );
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let expected_width = width_twips as f64 / 20.0;
        let expected_height = height_twips as f64 / 20.0;
        assert!(
            (page.size.width - expected_width).abs() < 1.0,
            "Expected width ~{expected_width}, got {}",
            page.size.width
        );
        assert!(
            (page.size.height - expected_height).abs() < 1.0,
            "Expected height ~{expected_height}, got {}",
            page.size.height
        );
    }

    #[test]
    fn test_custom_margins_extracted() {
        // Margins: 0.5 inch = 720 twips = 36pt
        let data = build_docx_bytes_with_page_setup(
            vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test"))],
            12240,
            15840,
            720,
            720,
            720,
            720,
        );
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let expected_margin = 720.0 / 20.0; // 36pt
        assert!(
            (page.margins.top - expected_margin).abs() < 1.0,
            "Expected top margin ~{expected_margin}, got {}",
            page.margins.top
        );
        assert!((page.margins.bottom - expected_margin).abs() < 1.0);
        assert!((page.margins.left - expected_margin).abs() < 1.0);
        assert!((page.margins.right - expected_margin).abs() < 1.0);
    }

    // ----- Error handling tests -----

    #[test]
    fn test_parse_invalid_data_returns_error() {
        let parser = DocxParser;
        let result = parser.parse(b"not a valid docx file", &ConvertOptions::default());
        assert!(result.is_err());
        match result.unwrap_err() {
            ConvertError::Parse(_) => {}
            other => panic!("Expected Parse error, got: {other:?}"),
        }
    }

    #[test]
    fn test_parse_error_includes_library_name() {
        let parser = DocxParser;
        let result = parser.parse(b"not a valid docx file", &ConvertOptions::default());
        let err = result.unwrap_err();
        let msg = err.to_string();
        assert!(
            msg.contains("docx-rs"),
            "Parse error should include upstream library name 'docx-rs', got: {msg}"
        );
    }

    // ----- Text style defaults -----

    #[test]
    fn test_parsed_runs_have_default_text_style() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Plain text")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        let run = &para.runs[0];
        // Plain text should have default style (all None)
        assert!(run.style.bold.is_none() || run.style.bold == Some(false));
        assert!(run.style.italic.is_none() || run.style.italic == Some(false));
        assert!(run.style.underline.is_none() || run.style.underline == Some(false));
    }

    #[test]
    fn test_parsed_paragraphs_have_default_style() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Test")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        // Default paragraph style should have no explicit alignment
        assert!(para.style.alignment.is_none());
    }

    // ----- Inline formatting tests (US-004) -----

    /// Helper: extract the first run from the first paragraph of a parsed document.
    fn first_run(doc: &Document) -> &Run {
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        &para.runs[0]
    }

    #[test]
    fn test_bold_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Bold text").bold()),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.bold, Some(true));
    }

    #[test]
    fn test_italic_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Italic text").italic()),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.italic, Some(true));
    }

    #[test]
    fn test_underline_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Underlined text")
                    .underline("single"),
            ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.underline, Some(true));
    }

    #[test]
    fn test_strikethrough_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Struck text").strike()),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.strikethrough, Some(true));
    }

    #[test]
    fn test_font_size_extracted() {
        // docx-rs size is in half-points: 24 half-points = 12pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Sized text").size(24)),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.font_size, Some(12.0));
    }

    #[test]
    fn test_letter_spacing_extracted() {
        // docx-rs character spacing is in twips: 40 twips = 2pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Tracked text")
                    .character_spacing(40),
            ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.letter_spacing, Some(2.0));
    }

    #[test]
    fn test_font_color_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Red text").color("FF0000")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.color, Some(Color::new(255, 0, 0)));
    }

    #[test]
    fn test_font_family_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Arial text")
                    .fonts(docx_rs::RunFonts::new().ascii("Arial")),
            ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.font_family, Some("Arial".to_string()));
    }

    #[test]
    fn test_combined_formatting_extracted() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Styled text")
                    .bold()
                    .italic()
                    .underline("single")
                    .strike()
                    .size(28) // 14pt
                    .color("0000FF")
                    .fonts(docx_rs::RunFonts::new().ascii("Courier")),
            ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);
        assert_eq!(run.style.bold, Some(true));
        assert_eq!(run.style.italic, Some(true));
        assert_eq!(run.style.underline, Some(true));
        assert_eq!(run.style.strikethrough, Some(true));
        assert_eq!(run.style.font_size, Some(14.0));
        assert_eq!(run.style.color, Some(Color::new(0, 0, 255)));
        assert_eq!(run.style.font_family, Some("Courier".to_string()));
    }

    #[test]
    fn test_plain_text_has_no_formatting() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Plain text")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);
        assert!(run.style.bold.is_none());
        assert!(run.style.italic.is_none());
        assert!(run.style.underline.is_none());
        assert!(run.style.strikethrough.is_none());
        assert!(run.style.font_size.is_none());
        assert!(run.style.letter_spacing.is_none());
        assert!(run.style.color.is_none());
        assert!(run.style.font_family.is_none());
    }

    // ----- Paragraph formatting tests (US-005) -----

    /// Helper: extract the first paragraph from a parsed document.
    fn first_paragraph(doc: &Document) -> &Paragraph {
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph block"),
        }
    }

    /// Helper: get all blocks from the first page.
    fn all_blocks(doc: &Document) -> &[Block] {
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        &page.content
    }

    #[test]
    fn test_paragraph_alignment_center() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Centered"))
                .align(docx_rs::AlignmentType::Center),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Center));
    }

    #[test]
    fn test_paragraph_alignment_right() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Right"))
                .align(docx_rs::AlignmentType::Right),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Right));
    }

    #[test]
    fn test_paragraph_alignment_left() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Left"))
                .align(docx_rs::AlignmentType::Left),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Left));
    }

    #[test]
    fn test_paragraph_alignment_justify() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Justified"))
                .align(docx_rs::AlignmentType::Both),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Justify));
    }

    #[test]
    fn test_paragraph_indent_left() {
        // 720 twips = 36pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Indented"))
                .indent(Some(720), None, None, None),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.indent_left, Some(36.0));
    }

    #[test]
    fn test_paragraph_indent_right() {
        // 360 twips = 18pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Indented"))
                .indent(None, None, Some(360), None),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.indent_right, Some(18.0));
    }

    #[test]
    fn test_paragraph_indent_first_line() {
        // first line indent: 480 twips = 24pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("First line indented"))
                .indent(
                    None,
                    Some(docx_rs::SpecialIndentType::FirstLine(480)),
                    None,
                    None,
                ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.indent_first_line, Some(24.0));
    }

    #[test]
    fn test_paragraph_indent_hanging() {
        // hanging indent: 360 twips = 18pt (negative first-line indent)
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Hanging indent"))
                .indent(
                    Some(720),
                    Some(docx_rs::SpecialIndentType::Hanging(360)),
                    None,
                    None,
                ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.indent_left, Some(36.0));
        assert_eq!(para.style.indent_first_line, Some(-18.0));
    }

    #[test]
    fn test_paragraph_line_spacing_auto() {
        // Auto line spacing: line=480 means 480/240 = 2.0 (double spacing)
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Double spaced"))
                .line_spacing(
                    docx_rs::LineSpacing::new()
                        .line_rule(docx_rs::LineSpacingType::Auto)
                        .line(480),
                ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        match para.style.line_spacing {
            Some(LineSpacing::Proportional(factor)) => {
                assert!(
                    (factor - 2.0).abs() < 0.01,
                    "Expected 2.0 (double spacing), got {factor}"
                );
            }
            other => panic!("Expected Proportional line spacing, got {other:?}"),
        }
    }

    #[test]
    fn test_paragraph_line_spacing_exact() {
        // Exact line spacing: line=240 twips = 12pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Exact spaced"))
                .line_spacing(
                    docx_rs::LineSpacing::new()
                        .line_rule(docx_rs::LineSpacingType::Exact)
                        .line(240),
                ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        match para.style.line_spacing {
            Some(LineSpacing::Exact(pts)) => {
                assert!((pts - 12.0).abs() < 0.01, "Expected 12pt, got {pts}");
            }
            other => panic!("Expected Exact line spacing, got {other:?}"),
        }
    }

    #[test]
    fn test_paragraph_space_before_after() {
        // before=240 twips = 12pt, after=120 twips = 6pt
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Spaced paragraph"))
                .line_spacing(docx_rs::LineSpacing::new().before(240).after(120)),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.space_before, Some(12.0));
        assert_eq!(para.style.space_after, Some(6.0));
    }

    #[test]
    fn test_paragraph_page_break_before() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Before break")),
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("After break"))
                .page_break_before(true),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let blocks = all_blocks(&doc);
        // Should have: Paragraph("Before break"), PageBreak, Paragraph("After break")
        assert_eq!(blocks.len(), 3, "Expected 3 blocks, got {}", blocks.len());
        assert!(matches!(&blocks[0], Block::Paragraph(_)));
        assert!(matches!(&blocks[1], Block::PageBreak));
        assert!(matches!(&blocks[2], Block::Paragraph(_)));
    }

    #[test]
    fn test_paragraph_combined_formatting() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Styled paragraph"))
                .align(docx_rs::AlignmentType::Center)
                .indent(
                    Some(720),
                    Some(docx_rs::SpecialIndentType::FirstLine(360)),
                    None,
                    None,
                )
                .line_spacing(
                    docx_rs::LineSpacing::new()
                        .line_rule(docx_rs::LineSpacingType::Auto)
                        .line(360)
                        .before(120)
                        .after(60),
                ),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(para.style.alignment, Some(Alignment::Center));
        assert_eq!(para.style.indent_left, Some(36.0));
        assert_eq!(para.style.indent_first_line, Some(18.0));
        assert_eq!(para.style.space_before, Some(6.0));
        assert_eq!(para.style.space_after, Some(3.0));
        match para.style.line_spacing {
            Some(LineSpacing::Proportional(factor)) => {
                assert!(
                    (factor - 1.5).abs() < 0.01,
                    "Expected 1.5 spacing, got {factor}"
                );
            }
            other => panic!("Expected Proportional line spacing, got {other:?}"),
        }
    }

    #[test]
    fn test_multiple_runs_with_different_formatting() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Bold ").bold())
                .add_run(docx_rs::Run::new().add_text("Italic ").italic())
                .add_run(docx_rs::Run::new().add_text("Plain")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };
        assert_eq!(para.runs.len(), 3);
        assert_eq!(para.runs[0].style.bold, Some(true));
        assert!(para.runs[0].style.italic.is_none());
        assert!(para.runs[1].style.bold.is_none());
        assert_eq!(para.runs[1].style.italic, Some(true));
        assert!(para.runs[2].style.bold.is_none());
        assert!(para.runs[2].style.italic.is_none());
    }

    // ----- Table parsing tests (US-007) -----

    /// Helper: build a DOCX with a table using docx-rs builder.
    fn build_docx_with_table(table: docx_rs::Table) -> Vec<u8> {
        let docx = docx_rs::Docx::new().add_table(table);
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: extract the first table block from a parsed document.
    fn first_table(doc: &Document) -> &crate::ir::Table {
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        for block in &page.content {
            if let Block::Table(t) = block {
                return t;
            }
        }
        panic!("No Table block found");
    }

    #[test]
    fn test_table_simple_2x2() {
        let table = docx_rs::Table::new(vec![
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A1")),
                ),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B1")),
                ),
            ]),
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A2")),
                ),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B2")),
                ),
            ]),
        ])
        .set_grid(vec![2000, 3000]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.rows.len(), 2);
        assert_eq!(t.rows[0].cells.len(), 2);
        assert_eq!(t.rows[1].cells.len(), 2);

        // Check cell content
        let cell_text = |row: usize, col: usize| -> String {
            t.rows[row].cells[col]
                .content
                .iter()
                .filter_map(|b| match b {
                    Block::Paragraph(p) => {
                        Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                    }
                    _ => None,
                })
                .collect::<String>()
        };
        assert_eq!(cell_text(0, 0), "A1");
        assert_eq!(cell_text(0, 1), "B1");
        assert_eq!(cell_text(1, 0), "A2");
        assert_eq!(cell_text(1, 1), "B2");
    }

    #[test]
    fn test_table_column_widths_from_grid() {
        // Grid widths in twips: 2000, 3000 → 100pt, 150pt
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A")),
            ),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B")),
            ),
        ])])
        .set_grid(vec![2000, 3000]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.column_widths.len(), 2);
        assert!(
            (t.column_widths[0] - 100.0).abs() < 0.1,
            "Expected 100pt, got {}",
            t.column_widths[0]
        );
        assert!(
            (t.column_widths[1] - 150.0).abs() < 0.1,
            "Expected 150pt, got {}",
            t.column_widths[1]
        );
    }

    #[test]
    fn test_table_column_widths_from_cell_widths_without_grid() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A")))
                .width(2000, docx_rs::WidthType::Dxa),
            docx_rs::TableCell::new()
                .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B")))
                .width(3000, docx_rs::WidthType::Dxa),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.column_widths.len(), 2);
        assert!(
            (t.column_widths[0] - 100.0).abs() < 0.1,
            "Expected 100pt, got {}",
            t.column_widths[0]
        );
        assert!(
            (t.column_widths[1] - 150.0).abs() < 0.1,
            "Expected 150pt, got {}",
            t.column_widths[1]
        );
    }

    #[test]
    fn test_table_column_widths_from_spanned_cell_widths_without_grid() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Merged")),
                )
                .grid_span(2)
                .width(4000, docx_rs::WidthType::Dxa),
            docx_rs::TableCell::new()
                .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C")))
                .width(2000, docx_rs::WidthType::Dxa),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.column_widths.len(), 3);
        assert!(
            (t.column_widths[0] - 100.0).abs() < 0.1,
            "Expected first merged column to be 100pt, got {}",
            t.column_widths[0]
        );
        assert!(
            (t.column_widths[1] - 100.0).abs() < 0.1,
            "Expected second merged column to be 100pt, got {}",
            t.column_widths[1]
        );
        assert!(
            (t.column_widths[2] - 100.0).abs() < 0.1,
            "Expected final column to be 100pt, got {}",
            t.column_widths[2]
        );
    }

    #[test]
    fn test_scan_table_headers_counts_only_leading_rows() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:p><w:r><w:t>D1</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>Ignored</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
        <w:tbl>
            <w:tr>
                <w:tc><w:p><w:r><w:t>Only body</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
    </w:body>
</w:document>"#;

        let headers = scan_table_headers(document_xml);

        assert_eq!(headers.len(), 2);
        assert_eq!(headers[0].repeat_rows, 2);
        assert_eq!(headers[1].repeat_rows, 0);
    }

    #[test]
    fn test_table_header_rows_from_raw_docx_xml() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:tbl>
            <w:tblPr/>
            <w:tblGrid>
                <w:gridCol w:w="2000"/>
                <w:gridCol w:w="2000"/>
            </w:tblGrid>
            <w:tr>
                <w:trPr><w:tblHeader/></w:trPr>
                <w:tc><w:p><w:r><w:t>Header A</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Header B</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:p><w:r><w:t>Body A</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>Body B</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.header_row_count, 1);
    }

    #[test]
    fn test_table_default_cell_margins_from_table_property() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cell")),
            ),
        ])])
        .margins(docx_rs::TableCellMargins::new().margin(40, 60, 20, 80));

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        assert_eq!(
            t.default_cell_padding,
            Some(Insets {
                top: 2.0,
                right: 3.0,
                bottom: 1.0,
                left: 4.0,
            })
        );
        assert!(t.rows[0].cells[0].padding.is_none());
    }

    #[test]
    fn test_table_cell_margins_override_table_defaults() {
        let mut cell = docx_rs::TableCell::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cell")));
        cell.property = docx_rs::TableCellProperty::new()
            .margin_top(100, docx_rs::WidthType::Dxa)
            .margin_left(120, docx_rs::WidthType::Dxa);

        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![cell])])
            .margins(docx_rs::TableCellMargins::new().margin(20, 40, 60, 80));

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        assert_eq!(
            t.default_cell_padding,
            Some(Insets {
                top: 1.0,
                right: 2.0,
                bottom: 3.0,
                left: 4.0,
            })
        );
        assert_eq!(
            t.rows[0].cells[0].padding,
            Some(Insets {
                top: 5.0,
                right: 2.0,
                bottom: 3.0,
                left: 6.0,
            })
        );
    }

    #[test]
    fn test_table_alignment_from_table_property() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Centered")),
            ),
        ])])
        .align(docx_rs::TableAlignmentType::Center);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.alignment, Some(Alignment::Center));
    }

    #[test]
    fn test_table_cell_with_formatted_text() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Bold").bold())
                    .add_run(docx_rs::Run::new().add_text(" and italic").italic()),
            ),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        let cell = &t.rows[0].cells[0];
        let para = match &cell.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph in cell"),
        };
        assert_eq!(para.runs.len(), 2);
        assert_eq!(para.runs[0].text, "Bold");
        assert_eq!(para.runs[0].style.bold, Some(true));
        assert_eq!(para.runs[1].text, " and italic");
        assert_eq!(para.runs[1].style.italic, Some(true));
    }

    #[test]
    fn test_table_colspan_via_grid_span() {
        let table = docx_rs::Table::new(vec![
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(
                        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Merged")),
                    )
                    .grid_span(2),
            ]),
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A2")),
                ),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B2")),
                ),
            ]),
        ])
        .set_grid(vec![2000, 2000]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        // First row: one merged cell with colspan=2
        assert_eq!(t.rows[0].cells.len(), 1);
        assert_eq!(t.rows[0].cells[0].col_span, 2);

        // Second row: two normal cells
        assert_eq!(t.rows[1].cells.len(), 2);
        assert_eq!(t.rows[1].cells[0].col_span, 1);
    }

    #[test]
    fn test_table_rowspan_via_vmerge() {
        let table = docx_rs::Table::new(vec![
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(
                        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Tall")),
                    )
                    .vertical_merge(docx_rs::VMergeType::Restart),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B1")),
                ),
            ]),
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(docx_rs::Paragraph::new())
                    .vertical_merge(docx_rs::VMergeType::Continue),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B2")),
                ),
            ]),
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(docx_rs::Paragraph::new())
                    .vertical_merge(docx_rs::VMergeType::Continue),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B3")),
                ),
            ]),
        ])
        .set_grid(vec![2000, 2000]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.rows.len(), 3);

        // First row: the restart cell should have rowspan=3
        let tall_cell = &t.rows[0].cells[0];
        assert_eq!(tall_cell.row_span, 3);

        // Second and third rows: continue cells should be skipped
        // so rows[1] and rows[2] should have only 1 cell each (B2, B3)
        assert_eq!(t.rows[1].cells.len(), 1);
        assert_eq!(t.rows[2].cells.len(), 1);
    }

    #[test]
    fn test_table_exact_row_height_and_cell_vertical_align() {
        let table = docx_rs::Table::new(vec![
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(
                        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Centered")),
                    )
                    .vertical_align(docx_rs::VAlignType::Center),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Peer")),
                ),
            ])
            .row_height(36.0)
            .height_rule(docx_rs::HeightRule::Exact),
        ])
        .set_grid(vec![2000, 2000]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.rows[0].height, Some(36.0));
        assert_eq!(
            t.rows[0].cells[0].vertical_align,
            Some(CellVerticalAlign::Center)
        );
    }

    #[test]
    fn test_table_cell_background_color() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Red bg")),
                )
                .shading(docx_rs::Shading::new().fill("FF0000")),
            docx_rs::TableCell::new().add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("No bg")),
            ),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.rows[0].cells[0].background, Some(Color::new(255, 0, 0)));
        assert!(t.rows[0].cells[1].background.is_none());
    }

    #[test]
    fn test_table_cell_borders() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Bordered")),
                )
                .set_border(
                    docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Top)
                        .size(16)
                        .color("FF0000"),
                )
                .set_border(
                    docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Bottom)
                        .size(8)
                        .color("0000FF"),
                ),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        let cell = &t.rows[0].cells[0];
        let border = cell.border.as_ref().expect("Expected cell border");

        // Top: size=16 eighths → 2pt, color=FF0000
        let top = border.top.as_ref().expect("Expected top border");
        assert!(
            (top.width - 2.0).abs() < 0.01,
            "Expected 2pt, got {}",
            top.width
        );
        assert_eq!(top.color, Color::new(255, 0, 0));

        // Bottom: size=8 eighths → 1pt, color=0000FF
        let bottom = border.bottom.as_ref().expect("Expected bottom border");
        assert!(
            (bottom.width - 1.0).abs() < 0.01,
            "Expected 1pt, got {}",
            bottom.width
        );
        assert_eq!(bottom.color, Color::new(0, 0, 255));
    }

    #[test]
    fn test_table_cell_border_styles() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new()
                        .add_run(docx_rs::Run::new().add_text("Styled borders")),
                )
                .set_border(
                    docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Top)
                        .size(16)
                        .color("000000")
                        .border_type(docx_rs::BorderType::Dashed),
                )
                .set_border(
                    docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Bottom)
                        .size(8)
                        .color("0000FF")
                        .border_type(docx_rs::BorderType::Dotted),
                )
                .set_border(
                    docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Left)
                        .size(12)
                        .color("FF0000")
                        .border_type(docx_rs::BorderType::DotDash),
                )
                .set_border(
                    docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Right)
                        .size(16)
                        .color("00FF00")
                        .border_type(docx_rs::BorderType::Double),
                ),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        let cell = &t.rows[0].cells[0];
        let border = cell.border.as_ref().expect("Expected cell border");

        // Top: dashed
        let top = border.top.as_ref().expect("Expected top border");
        assert_eq!(top.style, BorderLineStyle::Dashed, "Top should be dashed");

        // Bottom: dotted
        let bottom = border.bottom.as_ref().expect("Expected bottom border");
        assert_eq!(
            bottom.style,
            BorderLineStyle::Dotted,
            "Bottom should be dotted"
        );

        // Left: dashDot
        let left = border.left.as_ref().expect("Expected left border");
        assert_eq!(
            left.style,
            BorderLineStyle::DashDot,
            "Left should be dashDot"
        );

        // Right: double
        let right = border.right.as_ref().expect("Expected right border");
        assert_eq!(
            right.style,
            BorderLineStyle::Double,
            "Right should be double"
        );
    }

    #[test]
    fn test_table_cell_solid_border_default_style() {
        // Single (default) border type should map to Solid
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Solid")),
                )
                .set_border(
                    docx_rs::TableCellBorder::new(docx_rs::TableCellBorderPosition::Top)
                        .size(16)
                        .color("000000"),
                    // Default border_type is Single → should map to Solid
                ),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);
        let cell = &t.rows[0].cells[0];
        let border = cell.border.as_ref().expect("Expected cell border");
        let top = border.top.as_ref().expect("Expected top border");
        assert_eq!(top.style, BorderLineStyle::Solid, "Single → Solid");
    }

    #[test]
    fn test_table_cell_with_multiple_paragraphs() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Para 1")),
                )
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Para 2")),
                ),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        let cell = &t.rows[0].cells[0];
        let paras: Vec<&str> = cell
            .content
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) if !p.runs.is_empty() => Some(p.runs[0].text.as_str()),
                _ => None,
            })
            .collect();
        assert!(paras.contains(&"Para 1"), "Expected 'Para 1' in cell");
        assert!(paras.contains(&"Para 2"), "Expected 'Para 2' in cell");
    }

    #[test]
    fn test_table_with_paragraph_before_and_after() {
        let data = {
            let docx = docx_rs::Docx::new()
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Before")),
                )
                .add_table(docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
                    docx_rs::TableCell::new().add_paragraph(
                        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Cell")),
                    ),
                ])]))
                .add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("After")),
                );
            let buf = Vec::new();
            let mut cursor = Cursor::new(buf);
            docx.build().pack(&mut cursor).unwrap();
            cursor.into_inner()
        };

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let blocks = all_blocks(&doc);

        // Should have: Paragraph("Before"), Table, Paragraph("After")
        assert!(
            blocks.len() >= 3,
            "Expected at least 3 blocks, got {}",
            blocks.len()
        );
        assert!(matches!(&blocks[0], Block::Paragraph(_)));
        let has_table = blocks.iter().any(|b| matches!(b, Block::Table(_)));
        assert!(has_table, "Expected a Table block");
    }

    #[test]
    fn test_table_colspan_and_rowspan_combined() {
        // 3x3 table with top-left 2x2 merged
        let table = docx_rs::Table::new(vec![
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(
                        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Big")),
                    )
                    .grid_span(2)
                    .vertical_merge(docx_rs::VMergeType::Restart),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C1")),
                ),
            ]),
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new()
                    .add_paragraph(docx_rs::Paragraph::new())
                    .grid_span(2)
                    .vertical_merge(docx_rs::VMergeType::Continue),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C2")),
                ),
            ]),
            docx_rs::TableRow::new(vec![
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("A3")),
                ),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("B3")),
                ),
                docx_rs::TableCell::new().add_paragraph(
                    docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("C3")),
                ),
            ]),
        ])
        .set_grid(vec![2000, 2000, 2000]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        // First row: "Big" (colspan=2, rowspan=2) + "C1"
        let big_cell = &t.rows[0].cells[0];
        assert_eq!(big_cell.col_span, 2, "Expected colspan=2");
        assert_eq!(big_cell.row_span, 2, "Expected rowspan=2");

        // Second row: continue cell skipped, so only "C2"
        assert_eq!(t.rows[1].cells.len(), 1);

        // Third row: three normal cells
        assert_eq!(t.rows[2].cells.len(), 3);
    }

    #[test]
    fn test_table_empty_cells() {
        let table = docx_rs::Table::new(vec![docx_rs::TableRow::new(vec![
            docx_rs::TableCell::new().add_paragraph(docx_rs::Paragraph::new()),
            docx_rs::TableCell::new().add_paragraph(docx_rs::Paragraph::new()),
        ])]);

        let data = build_docx_with_table(table);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let t = first_table(&doc);

        assert_eq!(t.rows.len(), 1);
        assert_eq!(t.rows[0].cells.len(), 2);
        // Empty cells should still have content (possibly empty paragraphs)
        for cell in &t.rows[0].cells {
            assert_eq!(cell.col_span, 1);
            assert_eq!(cell.row_span, 1);
        }
    }

    // ── Image extraction tests ──────────────────────────────────────────

    /// Build a minimal valid 1×1 red pixel BMP image.
    /// BMP is trivially decodable by the `image` crate (no compression).
    fn make_test_bmp() -> Vec<u8> {
        let mut bmp = Vec::new();
        // BMP file header (14 bytes)
        bmp.extend_from_slice(b"BM"); // magic
        bmp.extend_from_slice(&58u32.to_le_bytes()); // file size
        bmp.extend_from_slice(&0u32.to_le_bytes()); // reserved
        bmp.extend_from_slice(&54u32.to_le_bytes()); // pixel data offset

        // BITMAPINFOHEADER (40 bytes)
        bmp.extend_from_slice(&40u32.to_le_bytes()); // header size
        bmp.extend_from_slice(&1i32.to_le_bytes()); // width
        bmp.extend_from_slice(&1i32.to_le_bytes()); // height
        bmp.extend_from_slice(&1u16.to_le_bytes()); // color planes
        bmp.extend_from_slice(&24u16.to_le_bytes()); // bits per pixel (RGB)
        bmp.extend_from_slice(&0u32.to_le_bytes()); // compression (none)
        bmp.extend_from_slice(&4u32.to_le_bytes()); // image size (row aligned to 4 bytes)
        bmp.extend_from_slice(&0u32.to_le_bytes()); // x pixels/meter
        bmp.extend_from_slice(&0u32.to_le_bytes()); // y pixels/meter
        bmp.extend_from_slice(&0u32.to_le_bytes()); // total colors
        bmp.extend_from_slice(&0u32.to_le_bytes()); // important colors

        // Pixel data: 1 pixel BGR (red = 00 00 FF) + 1 byte padding
        bmp.extend_from_slice(&[0x00, 0x00, 0xFF, 0x00]);

        bmp
    }

    /// Build a DOCX containing an inline image using Pic::new() which processes
    /// the image through the `image` crate, ensuring valid PNG output in the ZIP.
    fn build_docx_with_image(width_px: u32, height_px: u32) -> Vec<u8> {
        let bmp_data = make_test_bmp();
        // Use Pic::new() which decodes the BMP and re-encodes as PNG internally
        let pic = docx_rs::Pic::new(&bmp_data).size(width_px * 9525, height_px * 9525);
        let para_with_image = docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic));
        let docx = docx_rs::Docx::new().add_paragraph(para_with_image);
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Build a minimal DOCX with a custom `document.xml` and one image relationship.
    fn build_docx_with_custom_image_document(document_xml: &str) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::FileOptions::default();

        zip.start_file("[Content_Types].xml", options).unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bmp" ContentType="image/bmp"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
        )
        .unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

        zip.start_file("word/_rels/document.xml.rels", options)
            .unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.bmp"/>
</Relationships>"#,
        )
        .unwrap();

        zip.start_file("word/document.xml", options).unwrap();
        std::io::Write::write_all(&mut zip, document_xml.as_bytes()).unwrap();

        zip.start_file("word/media/image1.bmp", options).unwrap();
        std::io::Write::write_all(&mut zip, &make_test_bmp()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    /// Helper: find all Image blocks in a FlowPage.
    fn find_images(doc: &Document) -> Vec<&ImageData> {
        let page = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };
        page.content
            .iter()
            .filter_map(|b| match b {
                Block::Image(img) => Some(img),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn test_docx_image_inline_basic() {
        let data = build_docx_with_image(100, 80);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let images = find_images(&doc);
        assert_eq!(images.len(), 1, "Expected exactly one image block");
        assert!(!images[0].data.is_empty(), "Image data should not be empty");
    }

    #[test]
    fn test_docx_image_format_is_png() {
        let data = build_docx_with_image(50, 50);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let images = find_images(&doc);
        assert_eq!(
            images[0].format,
            ImageFormat::Png,
            "Image format should be PNG"
        );
    }

    #[test]
    fn test_docx_vml_shape_image_is_emitted() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <w:body>
        <w:p>
            <w:r>
                <w:pict>
                    <v:shape id="VMLImage1" style="width:72pt;height:36pt">
                        <v:imagedata r:id="rIdImage1"/>
                    </v:shape>
                </w:pict>
            </w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_custom_image_document(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let images = find_images(&doc);
        assert_eq!(images.len(), 1, "Expected one VML image");
        assert_eq!(images[0].format, ImageFormat::Png);
        assert_eq!(images[0].width, Some(72.0));
        assert_eq!(images[0].height, Some(36.0));
    }

    #[test]
    fn test_docx_image_dimensions() {
        // 100px × 80px → EMU: 100*9525=952500, 80*9525=762000
        // EMU to points: 952500/12700=75.0, 762000/12700=60.0
        let data = build_docx_with_image(100, 80);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let images = find_images(&doc);
        let img = images[0];

        let width = img.width.expect("Expected width");
        let height = img.height.expect("Expected height");

        assert!(
            (width - 75.0).abs() < 0.1,
            "Expected width ~75pt, got {width}"
        );
        assert!(
            (height - 60.0).abs() < 0.1,
            "Expected height ~60pt, got {height}"
        );
    }

    #[test]
    fn test_docx_image_with_text_paragraphs() {
        let bmp_data = make_test_bmp();
        let pic = docx_rs::Pic::new(&bmp_data);
        let docx = docx_rs::Docx::new()
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Before image")),
            )
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)))
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("After image")),
            );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };

        let has_image = page.content.iter().any(|b| matches!(b, Block::Image(_)));
        assert!(has_image, "Expected an image block in the content");

        let has_before = page.content.iter().any(|b| match b {
            Block::Paragraph(p) => p.runs.iter().any(|r| r.text.contains("Before")),
            _ => false,
        });
        assert!(has_before, "Expected 'Before image' text");

        let has_after = page.content.iter().any(|b| match b {
            Block::Paragraph(p) => p.runs.iter().any(|r| r.text.contains("After")),
            _ => false,
        });
        assert!(has_after, "Expected 'After image' text");
    }

    #[test]
    fn test_docx_multiple_images() {
        let bmp_data = make_test_bmp();
        let pic1 = docx_rs::Pic::new(&bmp_data).size(100 * 9525, 100 * 9525);
        let pic2 = docx_rs::Pic::new(&bmp_data).size(200 * 9525, 150 * 9525);
        let docx = docx_rs::Docx::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic1)))
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic2)));
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let images = find_images(&doc);
        assert!(
            images.len() >= 2,
            "Expected at least 2 images, got {}",
            images.len()
        );
    }

    #[test]
    fn test_docx_image_data_contains_png_header() {
        let data = build_docx_with_image(50, 50);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let images = find_images(&doc);
        let img_data = &images[0].data;

        // docx-rs converts all images to PNG; verify PNG magic bytes
        assert!(
            img_data.len() >= 8 && img_data[0..4] == [0x89, 0x50, 0x4E, 0x47],
            "Image data should start with PNG magic bytes"
        );
    }

    #[test]
    fn test_docx_no_images_produces_no_image_blocks() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Just text")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };

        let image_count = page
            .content
            .iter()
            .filter(|b| matches!(b, Block::Image(_)))
            .count();
        assert_eq!(image_count, 0, "Expected no image blocks");
    }

    #[test]
    fn test_docx_image_with_custom_emu_size() {
        // Create image with specific EMU size via .size() override
        // 200pt × 100pt → 200*12700=2540000, 100*12700=1270000
        let bmp_data = make_test_bmp();
        let pic = docx_rs::Pic::new(&bmp_data).size(2_540_000, 1_270_000);
        let docx = docx_rs::Docx::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let images = find_images(&doc);
        assert!(!images.is_empty(), "Expected at least one image");
        let img = images[0];

        let width = img.width.expect("Expected width");
        let height = img.height.expect("Expected height");

        assert!(
            (width - 200.0).abs() < 0.1,
            "Expected width ~200pt, got {width}"
        );
        assert!(
            (height - 100.0).abs() < 0.1,
            "Expected height ~100pt, got {height}"
        );
    }

    // ----- List parsing tests -----

    /// Helper: build a DOCX with numbering definitions and list paragraphs.
    fn build_docx_with_numbering(
        abstract_nums: Vec<docx_rs::AbstractNumbering>,
        numberings: Vec<docx_rs::Numbering>,
        paragraphs: Vec<docx_rs::Paragraph>,
    ) -> Vec<u8> {
        let mut nums = docx_rs::Numberings::new();
        for an in abstract_nums {
            nums = nums.add_abstract_numbering(an);
        }
        for n in numberings {
            nums = nums.add_numbering(n);
        }

        let mut docx = docx_rs::Docx::new().numberings(nums);
        for p in paragraphs {
            docx = docx.add_paragraph(p);
        }
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    #[test]
    fn test_parse_simple_bulleted_list() {
        // Create a bullet list: abstractNum with format "bullet", numId=1, ilvl=0
        let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("bullet"),
            docx_rs::LevelText::new("•"),
            docx_rs::LevelJc::new("left"),
        ));
        let numbering = docx_rs::Numbering::new(1, 0);

        let data = build_docx_with_numbering(
            vec![abstract_num],
            vec![numbering],
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Item A"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Item B"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Item C"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            ],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };

        // Should produce a single List block with 3 items
        let lists: Vec<&List> = page
            .content
            .iter()
            .filter_map(|b| match b {
                Block::List(l) => Some(l),
                _ => None,
            })
            .collect();
        assert_eq!(lists.len(), 1, "Expected 1 list block");
        assert_eq!(lists[0].kind, ListKind::Unordered);
        assert_eq!(lists[0].items.len(), 3);
        assert_eq!(lists[0].items[0].level, 0);
        assert_eq!(
            lists[0].level_styles.get(&0),
            Some(&ListLevelStyle {
                kind: ListKind::Unordered,
                numbering_pattern: None,
                full_numbering: false,
            })
        );

        // Verify item content
        let text0: String = lists[0].items[0]
            .content
            .iter()
            .flat_map(|p| p.runs.iter().map(|r| r.text.as_str()))
            .collect();
        assert_eq!(text0, "Item A");
    }

    #[test]
    fn test_parse_simple_numbered_list() {
        let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("decimal"),
            docx_rs::LevelText::new("%1."),
            docx_rs::LevelJc::new("left"),
        ));
        let numbering = docx_rs::Numbering::new(1, 0);

        let data = build_docx_with_numbering(
            vec![abstract_num],
            vec![numbering],
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("First"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Second"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            ],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };

        let lists: Vec<&List> = page
            .content
            .iter()
            .filter_map(|b| match b {
                Block::List(l) => Some(l),
                _ => None,
            })
            .collect();
        assert_eq!(lists.len(), 1, "Expected 1 list block");
        assert_eq!(lists[0].kind, ListKind::Ordered);
        assert_eq!(lists[0].items.len(), 2);
        assert_eq!(lists[0].items[0].start_at, Some(1));
        assert_eq!(
            lists[0].level_styles.get(&0),
            Some(&ListLevelStyle {
                kind: ListKind::Ordered,
                numbering_pattern: Some("1.".to_string()),
                full_numbering: false,
            })
        );
    }

    #[test]
    fn test_parse_nested_multi_level_list() {
        let abstract_num = docx_rs::AbstractNumbering::new(0)
            .add_level(docx_rs::Level::new(
                0,
                docx_rs::Start::new(1),
                docx_rs::NumberFormat::new("bullet"),
                docx_rs::LevelText::new("•"),
                docx_rs::LevelJc::new("left"),
            ))
            .add_level(docx_rs::Level::new(
                1,
                docx_rs::Start::new(1),
                docx_rs::NumberFormat::new("bullet"),
                docx_rs::LevelText::new("◦"),
                docx_rs::LevelJc::new("left"),
            ));
        let numbering = docx_rs::Numbering::new(1, 0);

        let data = build_docx_with_numbering(
            vec![abstract_num],
            vec![numbering],
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Top level"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Nested item"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(1)),
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Back to top"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            ],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };

        let lists: Vec<&List> = page
            .content
            .iter()
            .filter_map(|b| match b {
                Block::List(l) => Some(l),
                _ => None,
            })
            .collect();
        assert_eq!(lists.len(), 1, "Expected 1 list block");
        assert_eq!(lists[0].items.len(), 3);
        assert_eq!(lists[0].items[0].level, 0);
        assert_eq!(lists[0].items[1].level, 1);
        assert_eq!(lists[0].items[2].level, 0);
        assert_eq!(
            lists[0].level_styles.get(&1),
            Some(&ListLevelStyle {
                kind: ListKind::Unordered,
                numbering_pattern: None,
                full_numbering: false,
            })
        );
    }

    #[test]
    fn test_parse_numbered_list_start_override() {
        let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("decimal"),
            docx_rs::LevelText::new("%1."),
            docx_rs::LevelJc::new("left"),
        ));
        let numbering =
            docx_rs::Numbering::new(1, 0).add_override(docx_rs::LevelOverride::new(0).start(3));

        let data = build_docx_with_numbering(
            vec![abstract_num],
            vec![numbering],
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Third"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Fourth"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
            ],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let list = page
            .content
            .iter()
            .find_map(|block| match block {
                Block::List(list) => Some(list),
                _ => None,
            })
            .expect("Expected list block");

        assert_eq!(list.items[0].start_at, Some(3));
        assert_eq!(list.items[1].start_at, None);
        assert_eq!(
            list.level_styles.get(&0),
            Some(&ListLevelStyle {
                kind: ListKind::Ordered,
                numbering_pattern: Some("1.".to_string()),
                full_numbering: false,
            })
        );
    }

    #[test]
    fn test_parse_mixed_ordered_and_bulleted_levels() {
        let abstract_num = docx_rs::AbstractNumbering::new(0)
            .add_level(docx_rs::Level::new(
                0,
                docx_rs::Start::new(1),
                docx_rs::NumberFormat::new("decimal"),
                docx_rs::LevelText::new("%1."),
                docx_rs::LevelJc::new("left"),
            ))
            .add_level(docx_rs::Level::new(
                1,
                docx_rs::Start::new(1),
                docx_rs::NumberFormat::new("bullet"),
                docx_rs::LevelText::new("•"),
                docx_rs::LevelJc::new("left"),
            ));
        let numbering = docx_rs::Numbering::new(1, 0);

        let data = build_docx_with_numbering(
            vec![abstract_num],
            vec![numbering],
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Step"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Bullet child"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(1)),
            ],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let list = page
            .content
            .iter()
            .find_map(|block| match block {
                Block::List(list) => Some(list),
                _ => None,
            })
            .expect("Expected list block");

        assert_eq!(list.kind, ListKind::Ordered);
        assert_eq!(
            list.level_styles,
            BTreeMap::from([
                (
                    0,
                    ListLevelStyle {
                        kind: ListKind::Ordered,
                        numbering_pattern: Some("1.".to_string()),
                        full_numbering: false,
                    },
                ),
                (
                    1,
                    ListLevelStyle {
                        kind: ListKind::Unordered,
                        numbering_pattern: None,
                        full_numbering: false,
                    },
                ),
            ])
        );
    }

    #[test]
    fn test_parse_mixed_list_and_paragraphs() {
        // A list followed by a regular paragraph should produce two separate blocks
        let abstract_num = docx_rs::AbstractNumbering::new(0).add_level(docx_rs::Level::new(
            0,
            docx_rs::Start::new(1),
            docx_rs::NumberFormat::new("decimal"),
            docx_rs::LevelText::new("%1."),
            docx_rs::LevelJc::new("left"),
        ));
        let numbering = docx_rs::Numbering::new(1, 0);

        let data = build_docx_with_numbering(
            vec![abstract_num],
            vec![numbering],
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Item 1"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Item 2"))
                    .numbering(docx_rs::NumberingId::new(1), docx_rs::IndentLevel::new(0)),
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Regular paragraph")),
            ],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };

        // Should have at least a List block and a Paragraph block
        let list_count = page
            .content
            .iter()
            .filter(|b| matches!(b, Block::List(_)))
            .count();
        let para_count = page
            .content
            .iter()
            .filter(|b| matches!(b, Block::Paragraph(_)))
            .count();
        assert!(list_count >= 1, "Expected at least 1 list block");
        assert!(para_count >= 1, "Expected at least 1 paragraph block");
    }

    // ----- US-020: Header/footer parsing tests -----

    /// Helper: build a DOCX with a text header.
    fn build_docx_with_header(header_text: &str) -> Vec<u8> {
        let header = docx_rs::Header::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(header_text)),
        );
        let docx = docx_rs::Docx::new().header(header).add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
        );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: build a DOCX with a text footer.
    fn build_docx_with_footer(footer_text: &str) -> Vec<u8> {
        let footer = docx_rs::Footer::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(footer_text)),
        );
        let docx = docx_rs::Docx::new().footer(footer).add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
        );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: build a DOCX with a page number field in footer.
    fn build_docx_with_page_number_footer() -> Vec<u8> {
        let footer = docx_rs::Footer::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(
                docx_rs::Run::new()
                    .add_text("Page ")
                    .add_field_char(docx_rs::FieldCharType::Begin, false)
                    .add_instr_text(docx_rs::InstrText::PAGE(docx_rs::InstrPAGE::new()))
                    .add_field_char(docx_rs::FieldCharType::Separate, false)
                    .add_text("1")
                    .add_field_char(docx_rs::FieldCharType::End, false),
            ),
        );
        let docx = docx_rs::Docx::new().footer(footer).add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
        );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    #[test]
    fn test_parse_docx_with_text_header() {
        let data = build_docx_with_header("My Document Header");
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };

        // Should have a header
        assert!(page.header.is_some(), "FlowPage should have a header");
        let header = page.header.as_ref().unwrap();
        assert!(
            !header.paragraphs.is_empty(),
            "Header should have paragraphs"
        );

        // Find the text run in header
        let has_text = header.paragraphs.iter().any(|p| {
            p.elements.iter().any(|e| matches!(e, crate::ir::HFInline::Run(r) if r.text.contains("My Document Header")))
        });
        assert!(
            has_text,
            "Header should contain the text 'My Document Header'"
        );
    }

    #[test]
    fn test_parse_docx_with_text_footer() {
        let data = build_docx_with_footer("Footer Text");
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };

        assert!(page.footer.is_some(), "FlowPage should have a footer");
        let footer = page.footer.as_ref().unwrap();

        let has_text = footer.paragraphs.iter().any(|p| {
            p.elements
                .iter()
                .any(|e| matches!(e, crate::ir::HFInline::Run(r) if r.text.contains("Footer Text")))
        });
        assert!(has_text, "Footer should contain 'Footer Text'");
    }

    #[test]
    fn test_parse_docx_with_page_number_in_footer() {
        let data = build_docx_with_page_number_footer();
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };

        assert!(page.footer.is_some(), "Should have footer");
        let footer = page.footer.as_ref().unwrap();

        // Footer should contain a PageNumber element
        let has_page_num = footer.paragraphs.iter().any(|p| {
            p.elements
                .iter()
                .any(|e| matches!(e, crate::ir::HFInline::PageNumber))
        });
        assert!(has_page_num, "Footer should contain a PageNumber field");

        // Footer should also contain the "Page " text
        let has_text = footer.paragraphs.iter().any(|p| {
            p.elements
                .iter()
                .any(|e| matches!(e, crate::ir::HFInline::Run(r) if r.text.contains("Page ")))
        });
        assert!(
            has_text,
            "Footer should contain 'Page ' text before page number"
        );
    }

    /// Helper: build a DOCX with a total page count field in footer.
    fn build_docx_with_total_pages_footer() -> Vec<u8> {
        let footer = docx_rs::Footer::new().add_paragraph(
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Total "))
                .add_run(
                    docx_rs::Run::new()
                        .add_field_char(docx_rs::FieldCharType::Begin, false)
                        .add_instr_text(docx_rs::InstrText::NUMPAGES(docx_rs::InstrNUMPAGES::new()))
                        .add_field_char(docx_rs::FieldCharType::Separate, false)
                        .add_text("1")
                        .add_field_char(docx_rs::FieldCharType::End, false),
                ),
        );
        let docx = docx_rs::Docx::new()
            .footer(footer)
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body")));
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    #[test]
    fn test_parse_docx_with_total_pages_in_footer() {
        let data = build_docx_with_total_pages_footer();
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };

        let footer = page.footer.as_ref().expect("Should have footer");
        let has_total_pages = footer.paragraphs.iter().any(|p| {
            p.elements
                .iter()
                .any(|e| matches!(e, crate::ir::HFInline::TotalPages))
        });
        assert!(has_total_pages, "Footer should contain a TotalPages field");
    }

    #[test]
    fn test_parse_docx_multiple_sections_with_distinct_page_setup_and_headers() {
        let first_header = docx_rs::Header::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Section One Header")),
        );
        let second_header = docx_rs::Header::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Section Two Header")),
        );

        let first_section = docx_rs::Section::new()
            .page_size(docx_rs::PageSize::new().size(12240, 15840))
            .header(first_header)
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Section One")),
            );

        let docx = docx_rs::Docx::new()
            .add_section(first_section)
            .header(second_header)
            .page_size(15840, 12240)
            .add_paragraph(
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Section Two")),
            );
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(doc.pages.len(), 2, "Expected one FlowPage per DOCX section");

        let first_page = match &doc.pages[0] {
            Page::Flow(page) => page,
            _ => panic!("Expected first page to be FlowPage"),
        };
        let second_page = match &doc.pages[1] {
            Page::Flow(page) => page,
            _ => panic!("Expected second page to be FlowPage"),
        };

        assert!(
            (first_page.size.width - 612.0).abs() < 0.1,
            "first page width should come from first section"
        );
        assert!(
            (first_page.size.height - 792.0).abs() < 0.1,
            "first page height should come from first section"
        );
        assert!(
            (second_page.size.width - 792.0).abs() < 0.1,
            "second page width should come from final section"
        );
        assert!(
            (second_page.size.height - 612.0).abs() < 0.1,
            "second page height should come from final section"
        );

        let first_header_text = first_page
            .header
            .as_ref()
            .and_then(|hf| {
                hf.paragraphs
                    .iter()
                    .flat_map(|p| p.elements.iter())
                    .find_map(|e| match e {
                        crate::ir::HFInline::Run(run) => Some(run.text.as_str()),
                        _ => None,
                    })
            })
            .unwrap_or("");
        assert_eq!(first_header_text, "Section One Header");

        let second_header_text = second_page
            .header
            .as_ref()
            .and_then(|hf| {
                hf.paragraphs
                    .iter()
                    .flat_map(|p| p.elements.iter())
                    .find_map(|e| match e {
                        crate::ir::HFInline::Run(run) => Some(run.text.as_str()),
                        _ => None,
                    })
            })
            .unwrap_or("");
        assert_eq!(second_header_text, "Section Two Header");
    }

    #[test]
    fn test_parse_docx_with_header_and_footer() {
        let header = docx_rs::Header::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Header Text")),
        );
        let footer = docx_rs::Footer::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Footer Text")),
        );
        let docx = docx_rs::Docx::new()
            .header(header)
            .footer(footer)
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body")));
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };

        assert!(page.header.is_some(), "Should have header");
        assert!(page.footer.is_some(), "Should have footer");
    }

    #[test]
    fn test_parse_docx_without_header_footer() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Just text")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };

        assert!(page.header.is_none(), "No header expected");
        assert!(page.footer.is_none(), "No footer expected");
    }

    // ----- Page orientation tests -----

    #[test]
    fn test_portrait_document_width_less_than_height() {
        // Standard A4 portrait: 11906 x 16838 twips
        let data = build_docx_bytes_with_page_setup(
            vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Portrait"))],
            11906,
            16838,
            1440,
            1440,
            1440,
            1440,
        );
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        assert!(
            page.size.width < page.size.height,
            "Portrait: width ({}) should be < height ({})",
            page.size.width,
            page.size.height
        );
    }

    #[test]
    fn test_landscape_document_width_greater_than_height() {
        // Landscape A4: width and height swapped → 16838 x 11906 twips
        let data = build_docx_bytes_with_page_setup(
            vec![docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Landscape"))],
            16838,
            11906,
            1440,
            1440,
            1440,
            1440,
        );
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        assert!(
            page.size.width > page.size.height,
            "Landscape: width ({}) should be > height ({})",
            page.size.width,
            page.size.height
        );
        // Verify approximate values: 16838/20 = 841.9pt, 11906/20 = 595.3pt
        assert!(
            (page.size.width - 841.9).abs() < 1.0,
            "Expected width ~841.9, got {}",
            page.size.width
        );
        assert!(
            (page.size.height - 595.3).abs() < 1.0,
            "Expected height ~595.3, got {}",
            page.size.height
        );
    }

    #[test]
    fn test_default_document_is_portrait() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Default")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        // Default docx-rs page is A4 portrait
        assert!(
            page.size.width < page.size.height,
            "Default should be portrait: width ({}) < height ({})",
            page.size.width,
            page.size.height
        );
    }

    #[test]
    fn test_landscape_with_orient_attribute() {
        // Build a landscape DOCX using page_orient + swapped dimensions
        let mut docx = docx_rs::Docx::new()
            .page_size(16838, 11906)
            .page_orient(docx_rs::PageOrientationType::Landscape)
            .page_margin(
                docx_rs::PageMargin::new()
                    .top(1440)
                    .bottom(1440)
                    .left(1440)
                    .right(1440),
            );
        docx = docx.add_paragraph(
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Landscape with orient")),
        );
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        assert!(
            page.size.width > page.size.height,
            "Landscape with orient: width ({}) should be > height ({})",
            page.size.width,
            page.size.height
        );
    }

    #[test]
    fn test_extract_page_size_orient_landscape_swaps_dimensions() {
        // Edge case: orient=landscape but dimensions are portrait-style (w < h).
        // The parser should detect orient and swap width/height.
        let page_size = docx_rs::PageSize::new()
            .width(11906) // portrait w
            .height(16838) // portrait h
            .orient(docx_rs::PageOrientationType::Landscape);

        let result = extract_page_size(&page_size);
        assert!(
            result.width > result.height,
            "orient=landscape should ensure width ({}) > height ({})",
            result.width,
            result.height
        );
    }

    #[test]
    fn test_extract_page_size_no_orient_keeps_dimensions() {
        // No orient attribute: dimensions should be used as-is
        let page_size = docx_rs::PageSize::new().width(11906).height(16838);

        let result = extract_page_size(&page_size);
        // 11906/20 = 595.3, 16838/20 = 841.9
        assert!(
            result.width < result.height,
            "No orient: width ({}) should be < height ({})",
            result.width,
            result.height
        );
    }

    // ----- Document styles tests (US-022) -----

    /// Helper: build a DOCX with custom styles and paragraphs.
    fn build_docx_bytes_with_styles(
        paragraphs: Vec<docx_rs::Paragraph>,
        styles: Vec<docx_rs::Style>,
    ) -> Vec<u8> {
        let mut docx = docx_rs::Docx::new();
        for s in styles {
            docx = docx.add_style(s);
        }
        for p in paragraphs {
            docx = docx.add_paragraph(p);
        }
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    /// Helper: build a DOCX with an explicit stylesheet and paragraphs.
    fn build_docx_bytes_with_stylesheet(
        paragraphs: Vec<docx_rs::Paragraph>,
        styles: docx_rs::Styles,
    ) -> Vec<u8> {
        let mut docx = docx_rs::Docx::new().styles(styles);
        for p in paragraphs {
            docx = docx.add_paragraph(p);
        }
        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    #[test]
    fn test_heading1_style_applies_defaults() {
        // Create a Heading 1 style with outline level 0 (no explicit size/bold)
        let h1_style = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph)
            .name("Heading 1")
            .outline_lvl(0);

        let data = build_docx_bytes_with_styles(
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Title"))
                    .style("Heading1"),
            ],
            vec![h1_style],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);

        // Heading 1 default: 24pt bold
        assert_eq!(run.style.font_size, Some(24.0));
        assert_eq!(run.style.bold, Some(true));
    }

    #[test]
    fn test_heading2_style_applies_defaults() {
        let h2_style = docx_rs::Style::new("Heading2", docx_rs::StyleType::Paragraph)
            .name("Heading 2")
            .outline_lvl(1);

        let data = build_docx_bytes_with_styles(
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Subtitle"))
                    .style("Heading2"),
            ],
            vec![h2_style],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);

        // Heading 2 default: 20pt bold
        assert_eq!(run.style.font_size, Some(20.0));
        assert_eq!(run.style.bold, Some(true));
    }

    #[test]
    fn test_heading3_through_6_defaults() {
        // Test heading levels 3-6 with their expected default sizes
        let expected: Vec<(usize, &str, f64)> = vec![
            (2, "Heading3", 16.0), // H3
            (3, "Heading4", 14.0), // H4
            (4, "Heading5", 12.0), // H5
            (5, "Heading6", 11.0), // H6
        ];

        for (outline_lvl, style_id, expected_size) in expected {
            let style = docx_rs::Style::new(style_id, docx_rs::StyleType::Paragraph)
                .name(format!("Heading {}", outline_lvl + 1))
                .outline_lvl(outline_lvl);

            let data = build_docx_bytes_with_styles(
                vec![
                    docx_rs::Paragraph::new()
                        .add_run(docx_rs::Run::new().add_text("Heading text"))
                        .style(style_id),
                ],
                vec![style],
            );

            let parser = DocxParser;
            let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
            let run = first_run(&doc);

            assert_eq!(
                run.style.font_size,
                Some(expected_size),
                "Heading {} should have size {expected_size}pt",
                outline_lvl + 1
            );
            assert_eq!(
                run.style.bold,
                Some(true),
                "Heading {} should be bold",
                outline_lvl + 1
            );
        }
    }

    #[test]
    fn test_style_with_explicit_formatting() {
        // Style defines size=36 (half-points = 18pt) and bold
        let custom = docx_rs::Style::new("CustomStyle", docx_rs::StyleType::Paragraph)
            .name("Custom Style")
            .size(36) // 18pt in half-points
            .bold();

        let data = build_docx_bytes_with_styles(
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Custom styled"))
                    .style("CustomStyle"),
            ],
            vec![custom],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);

        assert_eq!(run.style.font_size, Some(18.0));
        assert_eq!(run.style.bold, Some(true));
    }

    #[test]
    fn test_explicit_run_formatting_overrides_style() {
        // Style says bold + 24pt (via heading defaults), but run explicitly sets size=20 (10pt)
        let h1_style = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph)
            .name("Heading 1")
            .outline_lvl(0);

        let data = build_docx_bytes_with_styles(
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Small heading").size(20)) // 10pt
                    .style("Heading1"),
            ],
            vec![h1_style],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);

        // Explicit size (10pt) overrides heading default (24pt)
        assert_eq!(run.style.font_size, Some(10.0));
        // Bold still comes from heading defaults since not explicitly overridden
        assert_eq!(run.style.bold, Some(true));
    }

    #[test]
    fn test_style_alignment_applied_to_paragraph() {
        let centered = docx_rs::Style::new("CenteredStyle", docx_rs::StyleType::Paragraph)
            .name("Centered")
            .align(docx_rs::AlignmentType::Center);

        let data = build_docx_bytes_with_styles(
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Centered paragraph"))
                    .style("CenteredStyle"),
            ],
            vec![centered],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);

        assert_eq!(para.style.alignment, Some(Alignment::Center));
    }

    #[test]
    fn test_normal_style_no_heading_defaults() {
        // Normal paragraphs (no heading) should not get heading defaults
        let normal = docx_rs::Style::new("Normal", docx_rs::StyleType::Paragraph).name("Normal");

        let data = build_docx_bytes_with_styles(
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Normal text"))
                    .style("Normal"),
            ],
            vec![normal],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);

        // Normal style should NOT have heading defaults
        assert!(run.style.font_size.is_none());
        assert!(run.style.bold.is_none());
    }

    #[test]
    fn test_heading_with_mixed_paragraphs() {
        // Document with Heading 1, Normal, Heading 2 paragraphs
        let h1 = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph)
            .name("Heading 1")
            .outline_lvl(0);
        let h2 = docx_rs::Style::new("Heading2", docx_rs::StyleType::Paragraph)
            .name("Heading 2")
            .outline_lvl(1);

        let data = build_docx_bytes_with_styles(
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Title"))
                    .style("Heading1"),
                docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Body text")),
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Subtitle"))
                    .style("Heading2"),
            ],
            vec![h1, h2],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let blocks = all_blocks(&doc);

        // First paragraph: Heading 1
        if let Block::Paragraph(p) = &blocks[0] {
            assert_eq!(p.runs[0].style.font_size, Some(24.0));
            assert_eq!(p.runs[0].style.bold, Some(true));
        } else {
            panic!("Expected Paragraph");
        }

        // Second paragraph: Normal (no style)
        if let Block::Paragraph(p) = &blocks[1] {
            assert!(p.runs[0].style.font_size.is_none());
            assert!(p.runs[0].style.bold.is_none());
        } else {
            panic!("Expected Paragraph");
        }

        // Third paragraph: Heading 2
        if let Block::Paragraph(p) = &blocks[2] {
            assert_eq!(p.runs[0].style.font_size, Some(20.0));
            assert_eq!(p.runs[0].style.bold, Some(true));
        } else {
            panic!("Expected Paragraph");
        }
    }

    #[test]
    fn test_style_with_color_and_font() {
        let custom = docx_rs::Style::new("Fancy", docx_rs::StyleType::Paragraph)
            .name("Fancy Style")
            .color("FF0000")
            .fonts(docx_rs::RunFonts::new().ascii("Georgia"));

        let data = build_docx_bytes_with_styles(
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Fancy text"))
                    .style("Fancy"),
            ],
            vec![custom],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let run = first_run(&doc);

        assert_eq!(run.style.color, Some(Color::new(255, 0, 0)));
        assert_eq!(run.style.font_family, Some("Georgia".to_string()));
    }

    #[test]
    fn test_runs_inherit_document_default_font() {
        let styles = docx_rs::Styles::new()
            .default_fonts(docx_rs::RunFonts::new().ascii("Raleway"))
            .default_size(18);

        let link = docx_rs::Hyperlink::new("https://example.com", docx_rs::HyperlinkType::External)
            .add_run(
                docx_rs::Run::new()
                    .color("1155cc")
                    .underline("single")
                    .add_text("Linked text"),
            );
        let paragraph = docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Plain text "))
            .add_hyperlink(link);
        let data = build_docx_bytes_with_stylesheet(vec![paragraph], styles);

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);

        assert_eq!(para.runs.len(), 2);
        assert_eq!(para.runs[0].style.font_family.as_deref(), Some("Raleway"));
        assert_eq!(para.runs[0].style.font_size, Some(9.0));
        assert_eq!(para.runs[1].href.as_deref(), Some("https://example.com"));
        assert_eq!(para.runs[1].style.font_family.as_deref(), Some("Raleway"));
        assert_eq!(para.runs[1].style.font_size, Some(9.0));
        assert_eq!(para.runs[1].style.color, Some(Color::new(17, 85, 204)));
        assert_eq!(para.runs[1].style.underline, Some(true));
    }

    // ----- Hyperlink tests (US-030) -----

    #[test]
    fn test_hyperlink_single_link_in_paragraph() {
        let link = docx_rs::Hyperlink::new("https://example.com", docx_rs::HyperlinkType::External)
            .add_run(docx_rs::Run::new().add_text("Click here"));
        let para = docx_rs::Paragraph::new().add_hyperlink(link);
        let data = build_docx_bytes(vec![para]);

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };

        assert_eq!(para.runs.len(), 1);
        assert_eq!(para.runs[0].text, "Click here");
        assert_eq!(para.runs[0].href, Some("https://example.com".to_string()));
    }

    #[test]
    fn test_hyperlink_mixed_text_and_link() {
        let link =
            docx_rs::Hyperlink::new("https://rust-lang.org", docx_rs::HyperlinkType::External)
                .add_run(docx_rs::Run::new().add_text("Rust"));
        let para = docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Visit "))
            .add_hyperlink(link)
            .add_run(docx_rs::Run::new().add_text(" for more."));
        let data = build_docx_bytes(vec![para]);

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };

        // Should have 3 runs: "Visit ", hyperlink "Rust", " for more."
        assert_eq!(para.runs.len(), 3);

        assert_eq!(para.runs[0].text, "Visit ");
        assert_eq!(para.runs[0].href, None);

        assert_eq!(para.runs[1].text, "Rust");
        assert_eq!(para.runs[1].href, Some("https://rust-lang.org".to_string()));

        assert_eq!(para.runs[2].text, " for more.");
        assert_eq!(para.runs[2].href, None);
    }

    #[test]
    fn test_hyperlink_multiple_links_in_paragraph() {
        let link1 = docx_rs::Hyperlink::new("https://first.com", docx_rs::HyperlinkType::External)
            .add_run(docx_rs::Run::new().add_text("First"));
        let link2 = docx_rs::Hyperlink::new("https://second.com", docx_rs::HyperlinkType::External)
            .add_run(docx_rs::Run::new().add_text("Second"));
        let para = docx_rs::Paragraph::new()
            .add_hyperlink(link1)
            .add_run(docx_rs::Run::new().add_text(" and "))
            .add_hyperlink(link2);
        let data = build_docx_bytes(vec![para]);

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(p) => p,
            _ => panic!("Expected FlowPage"),
        };
        let para = match &page.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected Paragraph"),
        };

        assert_eq!(para.runs.len(), 3);

        assert_eq!(para.runs[0].text, "First");
        assert_eq!(para.runs[0].href, Some("https://first.com".to_string()));

        assert_eq!(para.runs[1].text, " and ");
        assert_eq!(para.runs[1].href, None);

        assert_eq!(para.runs[2].text, "Second");
        assert_eq!(para.runs[2].href, Some("https://second.com".to_string()));
    }

    // ── Footnotes and endnotes ──────────────────────────────────────────

    #[test]
    fn test_footnote_single_in_paragraph() {
        let footnote = docx_rs::Footnote::new().add_content(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("This is a footnote.")),
        );

        let para = docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Some text"))
            .add_run(docx_rs::Run::new().add_footnote_reference(footnote))
            .add_run(docx_rs::Run::new().add_text(" after note."));

        let data = build_docx_bytes(vec![para]);

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected flow page"),
        };

        // Find the paragraph
        let para = match &flow.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected paragraph"),
        };

        // Should have runs including a footnote reference
        let note_run = para.runs.iter().find(|r| r.footnote.is_some());
        assert!(note_run.is_some(), "Expected a run with footnote content");
        assert_eq!(
            note_run.unwrap().footnote.as_deref(),
            Some("This is a footnote.")
        );
    }

    #[test]
    fn test_footnote_multiple_in_paragraph() {
        let fn1 = docx_rs::Footnote::new().add_content(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("First note.")),
        );
        let fn2 = docx_rs::Footnote::new().add_content(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Second note.")),
        );

        let para = docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("A"))
            .add_run(docx_rs::Run::new().add_footnote_reference(fn1))
            .add_run(docx_rs::Run::new().add_text(" B"))
            .add_run(docx_rs::Run::new().add_footnote_reference(fn2));

        let data = build_docx_bytes(vec![para]);

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected flow page"),
        };

        let para = match &flow.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected paragraph"),
        };

        let note_runs: Vec<_> = para.runs.iter().filter(|r| r.footnote.is_some()).collect();
        assert_eq!(note_runs.len(), 2);
        assert_eq!(note_runs[0].footnote.as_deref(), Some("First note."));
        assert_eq!(note_runs[1].footnote.as_deref(), Some("Second note."));
    }

    #[test]
    fn test_endnote_parsed_as_footnote() {
        // docx-rs doesn't support endnotes, so we build a minimal DOCX ZIP manually
        // with word/endnotes.xml and w:endnoteReference in document.xml
        let data = build_docx_with_endnote("Text before endnote", 1, "This is an endnote.");

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected flow page"),
        };

        let para = match &flow.content[0] {
            Block::Paragraph(p) => p,
            _ => panic!("Expected paragraph"),
        };

        let note_run = para.runs.iter().find(|r| r.footnote.is_some());
        assert!(note_run.is_some(), "Expected a run with endnote content");
        assert_eq!(
            note_run.unwrap().footnote.as_deref(),
            Some("This is an endnote.")
        );
    }

    /// Build a minimal DOCX ZIP with an endnote reference in the document body
    /// and endnote content in word/endnotes.xml.
    fn build_docx_with_endnote(text: &str, endnote_id: usize, endnote_text: &str) -> Vec<u8> {
        use std::io::Write;
        use zip::ZipWriter;
        use zip::write::FileOptions;

        let buf = Vec::new();
        let mut zip = ZipWriter::new(Cursor::new(buf));
        let opts = FileOptions::default();

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", opts).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>"#).unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", opts).unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

        // word/_rels/document.xml.rels
        zip.start_file("word/_rels/document.xml.rels", opts)
            .unwrap();
        zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>"#).unwrap();

        // word/document.xml - with endnoteReference
        zip.start_file("word/document.xml", opts).unwrap();
        let doc_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">{text}</w:t></w:r>
      <w:r>
        <w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>
        <w:endnoteReference w:id="{endnote_id}"/>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#
        );
        zip.write_all(doc_xml.as_bytes()).unwrap();

        // word/endnotes.xml
        zip.start_file("word/endnotes.xml", opts).unwrap();
        let endnotes_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="{endnote_id}">
    <w:p>
      <w:r><w:t xml:space="preserve">{endnote_text}</w:t></w:r>
    </w:p>
  </w:endnote>
</w:endnotes>"#
        );
        zip.write_all(endnotes_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    // ----- Table of Contents (TOC) parsing tests -----

    /// Helper: build a DOCX with a table of contents containing items.
    fn build_docx_with_toc(items: Vec<docx_rs::TableOfContentsItem>) -> Vec<u8> {
        let toc = items.into_iter().fold(
            docx_rs::TableOfContents::new()
                .heading_styles_range(1, 3)
                .alias("Table of contents"),
            |toc, item| toc.add_item(item),
        );

        let style1 =
            docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph).name("Heading 1");
        let style2 =
            docx_rs::Style::new("Heading2", docx_rs::StyleType::Paragraph).name("Heading 2");

        let p1 = docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Introduction"))
            .style("Heading1");
        let p2 = docx_rs::Paragraph::new()
            .add_run(docx_rs::Run::new().add_text("Details"))
            .style("Heading2");

        let docx = docx_rs::Docx::new()
            .add_style(style1)
            .add_style(style2)
            .add_table_of_contents(toc)
            .add_paragraph(p1)
            .add_paragraph(p2);

        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        cursor.into_inner()
    }

    #[test]
    fn test_docx_toc_with_entries() {
        let items = vec![
            docx_rs::TableOfContentsItem::new()
                .text("Introduction")
                .toc_key("_Toc00000000")
                .level(1)
                .page_ref("2"),
            docx_rs::TableOfContentsItem::new()
                .text("Details")
                .toc_key("_Toc00000001")
                .level(2)
                .page_ref("3"),
        ];

        let data = build_docx_with_toc(items);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        // The document should have content, and the TOC entry texts should be present
        let page = &doc.pages[0];
        let content = match page {
            Page::Flow(fp) => &fp.content,
            _ => panic!("Expected FlowPage"),
        };

        // Collect all text from all paragraphs
        let all_text: Vec<String> = content
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => {
                    let t: String = p.runs.iter().map(|r| r.text.clone()).collect();
                    if t.is_empty() { None } else { Some(t) }
                }
                _ => None,
            })
            .collect();

        // TOC entries "Introduction" and "Details" should appear in the output
        // (along with the heading paragraphs themselves)
        let has_introduction = all_text.iter().any(|t| t.contains("Introduction"));
        let has_details = all_text.iter().any(|t| t.contains("Details"));
        assert!(
            has_introduction,
            "Expected 'Introduction' in TOC output, got: {all_text:?}"
        );
        assert!(
            has_details,
            "Expected 'Details' in TOC output, got: {all_text:?}"
        );
    }

    #[test]
    fn test_docx_toc_multiple_entries_in_paragraph_list() {
        let items = vec![
            docx_rs::TableOfContentsItem::new()
                .text("Chapter One")
                .toc_key("_Toc10000001")
                .level(1)
                .page_ref("1"),
            docx_rs::TableOfContentsItem::new()
                .text("Chapter Two")
                .toc_key("_Toc10000002")
                .level(1)
                .page_ref("5"),
            docx_rs::TableOfContentsItem::new()
                .text("Section A")
                .toc_key("_Toc10000003")
                .level(2)
                .page_ref("10"),
        ];

        let data = build_docx_with_toc(items);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = &doc.pages[0];
        let content = match page {
            Page::Flow(fp) => &fp.content,
            _ => panic!("Expected FlowPage"),
        };

        // All three TOC entry texts should appear
        let all_text: Vec<String> = content
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => {
                    let t: String = p.runs.iter().map(|r| r.text.clone()).collect();
                    if t.is_empty() { None } else { Some(t) }
                }
                _ => None,
            })
            .collect();

        assert!(
            all_text.iter().any(|t| t.contains("Chapter One")),
            "Expected 'Chapter One' in output, got: {all_text:?}"
        );
        assert!(
            all_text.iter().any(|t| t.contains("Chapter Two")),
            "Expected 'Chapter Two' in output, got: {all_text:?}"
        );
        assert!(
            all_text.iter().any(|t| t.contains("Section A")),
            "Expected 'Section A' in output, got: {all_text:?}"
        );
    }

    #[test]
    fn test_docx_sdt_with_paragraphs() {
        // Test that generic SDTs with paragraph content are also parsed.
        // Build a DOCX manually using docx-rs StructuredDataTag.
        let sdt = docx_rs::StructuredDataTag::new().add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("SDT Content")),
        );

        let docx = docx_rs::Docx::new().add_structured_data_tag(sdt);

        let buf = Vec::new();
        let mut cursor = Cursor::new(buf);
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = &doc.pages[0];
        let content = match page {
            Page::Flow(fp) => &fp.content,
            _ => panic!("Expected FlowPage"),
        };

        let all_text: Vec<String> = content
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => {
                    let t: String = p.runs.iter().map(|r| r.text.clone()).collect();
                    if t.is_empty() { None } else { Some(t) }
                }
                _ => None,
            })
            .collect();

        assert!(
            all_text.iter().any(|t| t.contains("SDT Content")),
            "Expected 'SDT Content' in output, got: {all_text:?}"
        );
    }

    #[test]
    fn test_docx_drawing_text_box_paragraph_is_emitted() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p>
            <w:r><w:t>Before</w:t></w:r>
            <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:docPr id="1" name="Text Box 1"/>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p>
                                                <w:r><w:t>Inside box</w:t></w:r>
                                            </w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
            <w:r><w:t>After</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let texts: Vec<String> = match &doc.pages[0] {
            Page::Flow(flow) => flow
                .content
                .iter()
                .filter_map(|block| match block {
                    Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                    _ => None,
                })
                .collect(),
            _ => panic!("Expected FlowPage"),
        };

        assert_eq!(
            texts,
            vec![
                "Before".to_string(),
                "Inside box".to_string(),
                "After".to_string(),
            ]
        );
    }

    #[test]
    fn test_docx_drawing_text_box_multiple_paragraphs_are_emitted_in_order() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p><w:r><w:t>Lead-in</w:t></w:r></w:p>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:docPr id="1" name="Text Box 2"/>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p><w:r><w:t>First line</w:t></w:r></w:p>
                                            <w:p><w:r><w:t>Second line</w:t></w:r></w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
        <w:p><w:r><w:t>Tail</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let texts: Vec<String> = match &doc.pages[0] {
            Page::Flow(flow) => flow
                .content
                .iter()
                .filter_map(|block| match block {
                    Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                    _ => None,
                })
                .collect(),
            _ => panic!("Expected FlowPage"),
        };

        assert_eq!(
            texts,
            vec![
                "Lead-in".to_string(),
                "First line".to_string(),
                "Second line".to_string(),
                "Tail".to_string(),
            ]
        );
    }

    #[test]
    fn test_docx_drawing_text_box_table_is_emitted() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p><w:r><w:t>Before table box</w:t></w:r></w:p>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:docPr id="1" name="Text Box Table"/>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:tbl>
                                                <w:tblPr/>
                                                <w:tblGrid>
                                                    <w:gridCol w:w="2000"/>
                                                    <w:gridCol w:w="2000"/>
                                                </w:tblGrid>
                                                <w:tr>
                                                    <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
                                                    <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
                                                </w:tr>
                                            </w:tbl>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
        <w:p><w:r><w:t>After table box</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(flow) => flow,
            _ => panic!("Expected FlowPage"),
        };

        let has_table = flow
            .content
            .iter()
            .any(|block| matches!(block, Block::Table(_)));
        assert!(has_table, "Expected a table extracted from text box");

        let table = first_table(&doc);
        assert_eq!(table.rows.len(), 1);
        assert_eq!(table.rows[0].cells.len(), 2);

        let cell_text: Vec<String> = table.rows[0]
            .cells
            .iter()
            .map(|cell| {
                cell.content
                    .iter()
                    .filter_map(|block| match block {
                        Block::Paragraph(p) => Some(
                            p.runs
                                .iter()
                                .map(|run| run.text.as_str())
                                .collect::<String>(),
                        ),
                        _ => None,
                    })
                    .collect::<String>()
            })
            .collect();
        assert_eq!(cell_text, vec!["A".to_string(), "B".to_string()]);
    }

    #[test]
    fn test_docx_vml_text_box_paragraph_is_emitted() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml">
    <w:body>
        <w:p>
            <w:r><w:t>Before</w:t></w:r>
            <w:r>
                <w:pict>
                    <v:shape id="TextBox1" style="width:100pt;height:40pt">
                        <v:textbox>
                            <w:txbxContent>
                                <w:p><w:r><w:t>VML box</w:t></w:r></w:p>
                            </w:txbxContent>
                        </v:textbox>
                    </v:shape>
                </w:pict>
            </w:r>
            <w:r><w:t>After</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let texts: Vec<String> = match &doc.pages[0] {
            Page::Flow(flow) => flow
                .content
                .iter()
                .filter_map(|block| match block {
                    Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                    _ => None,
                })
                .collect(),
            _ => panic!("Expected FlowPage"),
        };

        assert_eq!(
            texts,
            vec![
                "Before".to_string(),
                "VML box".to_string(),
                "After".to_string(),
            ]
        );
    }

    #[test]
    fn test_docx_vml_text_box_multiple_paragraphs_are_emitted_in_order() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml">
    <w:body>
        <w:p><w:r><w:t>Lead-in</w:t></w:r></w:p>
        <w:p>
            <w:r>
                <w:pict>
                    <v:shape id="TextBox2" style="width:120pt;height:60pt">
                        <v:textbox>
                            <w:txbxContent>
                                <w:p><w:r><w:t>First VML line</w:t></w:r></w:p>
                                <w:p><w:r><w:t>Second VML line</w:t></w:r></w:p>
                            </w:txbxContent>
                        </v:textbox>
                    </v:shape>
                </w:pict>
            </w:r>
        </w:p>
        <w:p><w:r><w:t>Tail</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let texts: Vec<String> = match &doc.pages[0] {
            Page::Flow(flow) => flow
                .content
                .iter()
                .filter_map(|block| match block {
                    Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                    _ => None,
                })
                .collect(),
            _ => panic!("Expected FlowPage"),
        };

        assert_eq!(
            texts,
            vec![
                "Lead-in".to_string(),
                "First VML line".to_string(),
                "Second VML line".to_string(),
                "Tail".to_string(),
            ]
        );
    }

    #[test]
    fn test_docx_vml_floating_text_box_square_wrap() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:w10="urn:schemas-microsoft-com:office:word">
    <w:body>
        <w:p>
            <w:r><w:t>Before</w:t></w:r>
            <w:r>
                <w:pict>
                    <v:shape id="TextBox3"
                             style="position:absolute;margin-left:72pt;margin-top:36pt;width:144pt;height:72pt;z-index:1;visibility:visible;mso-wrap-style:square">
                        <v:textbox>
                            <w:txbxContent>
                                <w:p><w:r><w:t>VML floating box</w:t></w:r></w:p>
                            </w:txbxContent>
                        </v:textbox>
                    </v:shape>
                    <w10:wrap type="square"/>
                </w:pict>
            </w:r>
            <w:r><w:t>After</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let floating = find_floating_text_boxes(&doc);
        assert_eq!(floating.len(), 1, "Expected one floating VML text box");

        let ftb = floating[0];
        assert_eq!(ftb.wrap_mode, WrapMode::Square);
        assert!((ftb.offset_x - 72.0).abs() < 0.5);
        assert!((ftb.offset_y - 36.0).abs() < 0.5);
        assert!((ftb.width - 144.0).abs() < 0.5);
        assert!((ftb.height - 72.0).abs() < 0.5);

        let texts: Vec<String> = ftb
            .content
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                _ => None,
            })
            .collect();
        assert_eq!(texts, vec!["VML floating box".to_string()]);
    }

    #[test]
    fn test_docx_vml_floating_text_box_none_wrap() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:w10="urn:schemas-microsoft-com:office:word">
    <w:body>
        <w:p>
            <w:r>
                <w:pict>
                    <v:shape id="TextBox4"
                             style="position:absolute;margin-left:12pt;margin-top:18pt;width:90pt;height:40pt;z-index:1;visibility:visible;mso-wrap-style:square">
                        <v:textbox>
                            <w:txbxContent>
                                <w:p><w:r><w:t>No wrap box</w:t></w:r></w:p>
                            </w:txbxContent>
                        </v:textbox>
                    </v:shape>
                    <w10:wrap type="none"/>
                </w:pict>
            </w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let floating = find_floating_text_boxes(&doc);
        assert_eq!(floating.len(), 1, "Expected one floating VML text box");
        assert_eq!(floating[0].wrap_mode, WrapMode::None);
    }

    /// Helper: find all FloatingTextBox blocks in a FlowPage.
    fn find_floating_text_boxes(doc: &Document) -> Vec<&FloatingTextBox> {
        let page = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };
        page.content
            .iter()
            .filter_map(|b| match b {
                Block::FloatingTextBox(ftb) => Some(ftb),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn test_docx_floating_text_box_square_wrap() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p>
            <w:r><w:t>Before</w:t></w:r>
            <w:r>
                <w:drawing>
                    <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" allowOverlap="0" behindDoc="0" locked="0" layoutInCell="1" relativeHeight="251659264">
                        <wp:simplePos x="0" y="0"/>
                        <wp:positionH relativeFrom="margin"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                        <wp:positionV relativeFrom="margin"><wp:posOffset>457200</wp:posOffset></wp:positionV>
                        <wp:extent cx="1828800" cy="914400"/>
                        <wp:effectExtent l="0" t="0" r="0" b="0"/>
                        <wp:wrapSquare wrapText="bothSides"/>
                        <wp:docPr id="1" name="Anchored Text Box"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p><w:r><w:t>Inside anchored box</w:t></w:r></w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:anchor>
                </w:drawing>
            </w:r>
            <w:r><w:t>After</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let floating = find_floating_text_boxes(&doc);
        assert_eq!(floating.len(), 1, "Expected one floating text box");

        let ftb = floating[0];
        assert_eq!(ftb.wrap_mode, WrapMode::Square);
        assert!(
            (ftb.offset_x - 72.0).abs() < 0.5,
            "Expected offset_x ~72pt, got {}",
            ftb.offset_x
        );
        assert!(
            (ftb.offset_y - 36.0).abs() < 0.5,
            "Expected offset_y ~36pt, got {}",
            ftb.offset_y
        );
        assert!(
            (ftb.width - 144.0).abs() < 0.5,
            "Expected width ~144pt, got {}",
            ftb.width
        );
        assert!(
            (ftb.height - 72.0).abs() < 0.5,
            "Expected height ~72pt, got {}",
            ftb.height
        );

        let texts: Vec<String> = ftb
            .content
            .iter()
            .filter_map(|block| match block {
                Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect()),
                _ => None,
            })
            .collect();
        assert_eq!(texts, vec!["Inside anchored box".to_string()]);
    }

    #[test]
    fn test_docx_floating_text_box_top_and_bottom_wrap() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            mc:Ignorable="wps">
    <w:body>
        <w:p>
            <w:r>
                <w:drawing>
                    <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" allowOverlap="1" behindDoc="0" locked="0" layoutInCell="1" relativeHeight="251659264">
                        <wp:simplePos x="0" y="0"/>
                        <wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>
                        <wp:positionV relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionV>
                        <wp:extent cx="1270000" cy="635000"/>
                        <wp:effectExtent l="0" t="0" r="0" b="0"/>
                        <wp:wrapTopAndBottom/>
                        <wp:docPr id="2" name="Top Bottom Text Box"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                <wps:wsp>
                                    <wps:txbx>
                                        <w:txbxContent>
                                            <w:p><w:r><w:t>Top and bottom box</w:t></w:r></w:p>
                                        </w:txbxContent>
                                    </wps:txbx>
                                    <wps:bodyPr/>
                                </wps:wsp>
                            </a:graphicData>
                        </a:graphic>
                    </wp:anchor>
                </w:drawing>
            </w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let floating = find_floating_text_boxes(&doc);
        assert_eq!(floating.len(), 1, "Expected one floating text box");
        assert_eq!(floating[0].wrap_mode, WrapMode::TopAndBottom);
    }

    // ── Floating image (anchor) tests ──

    /// Helper: find all FloatingImage blocks in a FlowPage.
    fn find_floating_images(doc: &Document) -> Vec<&FloatingImage> {
        let page = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };
        page.content
            .iter()
            .filter_map(|b| match b {
                Block::FloatingImage(fi) => Some(fi),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn test_docx_floating_image_square_wrap() {
        // Build a floating image with wrapSquare (allow_overlap = false, floating)
        let bmp_data = make_test_bmp();
        let pic = docx_rs::Pic::new(&bmp_data)
            .size(2_540_000, 1_270_000) // 200pt × 100pt
            .floating()
            .offset_x(914_400) // 72pt (1 inch)
            .offset_y(457_200); // 36pt (0.5 inch)
        // allow_overlap defaults to false for floating() → docx-rs writes wrapSquare

        let docx = docx_rs::Docx::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let floating = find_floating_images(&doc);
        assert_eq!(floating.len(), 1, "Expected one floating image");

        let fi = floating[0];
        assert_eq!(fi.wrap_mode, WrapMode::Square);
        assert!(!fi.image.data.is_empty(), "Image data should not be empty");

        // Check dimensions
        let width = fi.image.width.expect("Expected width");
        let height = fi.image.height.expect("Expected height");
        assert!(
            (width - 200.0).abs() < 0.5,
            "Expected width ~200pt, got {width}"
        );
        assert!(
            (height - 100.0).abs() < 0.5,
            "Expected height ~100pt, got {height}"
        );
    }

    #[test]
    fn test_docx_floating_image_top_and_bottom_wrap() {
        // Build a DOCX manually with wrapTopAndBottom
        let bmp_data = make_test_bmp();
        let pic = docx_rs::Pic::new(&bmp_data)
            .size(1_270_000, 1_270_000) // 100pt × 100pt
            .floating()
            .overlapping(); // allow_overlap=true → docx-rs writes wrapNone

        // Build DOCX bytes
        let docx = docx_rs::Docx::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let mut data = cursor.into_inner();

        // Patch the DOCX XML to replace wrapNone with wrapTopAndBottom
        data = patch_docx_wrap_type(&data, "wp:wrapNone", "wp:wrapTopAndBottom");

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let floating = find_floating_images(&doc);
        assert_eq!(floating.len(), 1, "Expected one floating image");
        assert_eq!(floating[0].wrap_mode, WrapMode::TopAndBottom);
    }

    #[test]
    fn test_docx_floating_image_behind_wrap() {
        let bmp_data = make_test_bmp();
        let pic = docx_rs::Pic::new(&bmp_data)
            .size(1_270_000, 1_270_000)
            .floating()
            .overlapping(); // generates wrapNone

        let docx = docx_rs::Docx::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let mut data = cursor.into_inner();

        // Patch to wrapNone → behindDoc attribute + wrapNone
        // Behind text is indicated by behindDoc="1" attribute on wp:anchor,
        // combined with wrapNone. Our scan should detect this.
        data = patch_docx_behind_doc(&data);

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let floating = find_floating_images(&doc);
        assert_eq!(floating.len(), 1, "Expected one floating image");
        // behindDoc with wrapNone → Behind wrap mode
        assert_eq!(floating[0].wrap_mode, WrapMode::Behind);
    }

    #[test]
    fn test_docx_floating_image_position_offset() {
        let bmp_data = make_test_bmp();
        let pic = docx_rs::Pic::new(&bmp_data)
            .size(1_270_000, 1_270_000) // 100pt × 100pt
            .floating()
            .offset_x(914_400) // 72pt (1 inch in EMU)
            .offset_y(457_200); // 36pt (0.5 inch in EMU)

        let docx = docx_rs::Docx::new()
            .add_paragraph(docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_image(pic)));
        let mut cursor = Cursor::new(Vec::new());
        docx.build().pack(&mut cursor).unwrap();
        let data = cursor.into_inner();

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let floating = find_floating_images(&doc);
        assert_eq!(floating.len(), 1, "Expected one floating image");

        let fi = floating[0];
        assert!(
            (fi.offset_x - 72.0).abs() < 0.5,
            "Expected offset_x ~72pt, got {}",
            fi.offset_x
        );
        assert!(
            (fi.offset_y - 36.0).abs() < 0.5,
            "Expected offset_y ~36pt, got {}",
            fi.offset_y
        );
    }

    #[test]
    fn test_docx_inline_image_not_floating() {
        // Inline images should NOT become FloatingImage
        let data = build_docx_with_image(100, 80);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let floating = find_floating_images(&doc);
        assert_eq!(
            floating.len(),
            0,
            "Inline images should not be floating images"
        );

        let images = find_images(&doc);
        assert_eq!(images.len(), 1, "Should still find the inline image");
    }

    /// Helper: Patch a DOCX ZIP by replacing a wrap element in document.xml.
    fn patch_docx_wrap_type(data: &[u8], old_wrap: &str, new_wrap: &str) -> Vec<u8> {
        let mut archive = zip::ZipArchive::new(Cursor::new(data)).unwrap();
        let mut new_zip = zip::ZipWriter::new(Cursor::new(Vec::new()));

        for i in 0..archive.len() {
            let mut file = archive.by_index(i).unwrap();
            let name = file.name().to_string();
            let options = zip::write::FileOptions::default();
            new_zip.start_file(&name, options).unwrap();

            let mut contents = Vec::new();
            file.read_to_end(&mut contents).unwrap();

            if name == "word/document.xml" {
                let xml = String::from_utf8(contents).unwrap();
                // Replace self-closing: <wp:wrapNone /> → <wp:wrapTopAndBottom />
                let xml = xml
                    .replace(&format!("<{old_wrap} />"), &format!("<{new_wrap} />"))
                    .replace(&format!("<{old_wrap}/>"), &format!("<{new_wrap}/>"));
                std::io::Write::write_all(&mut new_zip, xml.as_bytes()).unwrap();
            } else {
                std::io::Write::write_all(&mut new_zip, &contents).unwrap();
            }
        }

        new_zip.finish().unwrap().into_inner()
    }

    /// Helper: Patch a DOCX ZIP to set behindDoc="1" on wp:anchor in document.xml.
    fn patch_docx_behind_doc(data: &[u8]) -> Vec<u8> {
        let mut archive = zip::ZipArchive::new(Cursor::new(data)).unwrap();
        let mut new_zip = zip::ZipWriter::new(Cursor::new(Vec::new()));

        for i in 0..archive.len() {
            let mut file = archive.by_index(i).unwrap();
            let name = file.name().to_string();
            let options = zip::write::FileOptions::default();
            new_zip.start_file(&name, options).unwrap();

            let mut contents = Vec::new();
            file.read_to_end(&mut contents).unwrap();

            if name == "word/document.xml" {
                let xml = String::from_utf8(contents).unwrap();
                // Replace existing behindDoc="0" with behindDoc="1"
                let xml = xml.replace("behindDoc=\"0\"", "behindDoc=\"1\"");
                std::io::Write::write_all(&mut new_zip, xml.as_bytes()).unwrap();
            } else {
                std::io::Write::write_all(&mut new_zip, &contents).unwrap();
            }
        }

        new_zip.finish().unwrap().into_inner()
    }

    // ── OMML math equation tests ──

    /// Build a DOCX ZIP with a custom document.xml containing OMML math.
    fn build_docx_with_math(document_xml: &str) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::FileOptions::default();

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", options).unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
        )
        .unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", options).unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

        // word/_rels/document.xml.rels
        zip.start_file("word/_rels/document.xml.rels", options)
            .unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        )
        .unwrap();

        // word/document.xml
        zip.start_file("word/document.xml", options).unwrap();
        std::io::Write::write_all(&mut zip, document_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn test_parse_docx_with_display_math_fraction() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:body>
        <w:p>
            <w:r><w:t>Before math</w:t></w:r>
        </w:p>
        <w:p>
            <m:oMathPara>
                <m:oMath>
                    <m:f>
                        <m:num><m:r><m:t>a</m:t></m:r></m:num>
                        <m:den><m:r><m:t>b</m:t></m:r></m:den>
                    </m:f>
                </m:oMath>
            </m:oMathPara>
        </w:p>
        <w:p>
            <w:r><w:t>After math</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_math(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(fp) => fp,
            _ => panic!("Expected FlowPage"),
        };

        // Should find a MathEquation block
        let math_blocks: Vec<&MathEquation> = page
            .content
            .iter()
            .filter_map(|b| match b {
                Block::MathEquation(m) => Some(m),
                _ => None,
            })
            .collect();

        assert!(
            !math_blocks.is_empty(),
            "Expected at least one MathEquation block, found none"
        );
        assert_eq!(math_blocks[0].content, "frac(a, b)");
        assert!(math_blocks[0].display);
    }

    #[test]
    fn test_parse_docx_with_inline_math_superscript() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:body>
        <w:p>
            <w:r><w:t>The value of </w:t></w:r>
            <m:oMath>
                <m:sSup>
                    <m:e><m:r><m:t>x</m:t></m:r></m:e>
                    <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                </m:sSup>
            </m:oMath>
            <w:r><w:t> is positive</w:t></w:r>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_math(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(fp) => fp,
            _ => panic!("Expected FlowPage"),
        };

        let math_blocks: Vec<&MathEquation> = page
            .content
            .iter()
            .filter_map(|b| match b {
                Block::MathEquation(m) => Some(m),
                _ => None,
            })
            .collect();

        assert!(
            !math_blocks.is_empty(),
            "Expected at least one MathEquation block"
        );
        assert_eq!(math_blocks[0].content, "x^2");
        assert!(!math_blocks[0].display);
    }

    #[test]
    fn test_parse_docx_with_complex_math() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:body>
        <w:p>
            <m:oMathPara>
                <m:oMath>
                    <m:r><m:t>E</m:t></m:r>
                    <m:r><m:t>=</m:t></m:r>
                    <m:r><m:t>m</m:t></m:r>
                    <m:sSup>
                        <m:e><m:r><m:t>c</m:t></m:r></m:e>
                        <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                    </m:sSup>
                </m:oMath>
            </m:oMathPara>
        </w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#;

        let data = build_docx_with_math(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let page = match &doc.pages[0] {
            Page::Flow(fp) => fp,
            _ => panic!("Expected FlowPage"),
        };

        let math_blocks: Vec<&MathEquation> = page
            .content
            .iter()
            .filter_map(|b| match b {
                Block::MathEquation(m) => Some(m),
                _ => None,
            })
            .collect();

        assert!(!math_blocks.is_empty());
        // Space before sSup separates run "m" from base "c" to prevent identifier
        // concatenation (both are semantically equivalent in Typst math: m × c²)
        assert_eq!(math_blocks[0].content, "E=m c^2");
        assert!(math_blocks[0].display);
    }

    /// Build a DOCX ZIP with a chart embedded in it.
    fn build_docx_with_chart(document_xml: &str, chart_xml: &str) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::FileOptions::default();

        // [Content_Types].xml
        zip.start_file("[Content_Types].xml", options).unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>"#,
        )
        .unwrap();

        // _rels/.rels
        zip.start_file("_rels/.rels", options).unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

        // word/_rels/document.xml.rels
        zip.start_file("word/_rels/document.xml.rels", options)
            .unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
</Relationships>"#,
        )
        .unwrap();

        // word/document.xml
        zip.start_file("word/document.xml", options).unwrap();
        std::io::Write::write_all(&mut zip, document_xml.as_bytes()).unwrap();

        // word/charts/chart1.xml
        zip.start_file("word/charts/chart1.xml", options).unwrap();
        std::io::Write::write_all(&mut zip, chart_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn test_parse_docx_with_bar_chart() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rId4"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let chart_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:strCache>
            <c:pt idx="0"><c:v>Q1</c:v></c:pt>
            <c:pt idx="1"><c:v>Q2</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache>
            <c:pt idx="0"><c:v>100</c:v></c:pt>
            <c:pt idx="1"><c:v>200</c:v></c:pt>
          </c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let data = build_docx_with_chart(document_xml, chart_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser
            .parse(&data, &ConvertOptions::default())
            .expect("parse should succeed");

        let content = match &doc.pages[0] {
            crate::ir::Page::Flow(fp) => &fp.content,
            _ => panic!("Expected FlowPage"),
        };
        let chart_blocks: Vec<&crate::ir::Chart> = content
            .iter()
            .filter_map(|b| match b {
                Block::Chart(c) => Some(c),
                _ => None,
            })
            .collect();

        assert_eq!(chart_blocks.len(), 1);
        assert_eq!(chart_blocks[0].chart_type, crate::ir::ChartType::Bar);
        assert_eq!(chart_blocks[0].title.as_deref(), Some("Sales"));
        assert_eq!(chart_blocks[0].categories, vec!["Q1", "Q2"]);
        assert_eq!(chart_blocks[0].series.len(), 1);
        assert_eq!(chart_blocks[0].series[0].name.as_deref(), Some("Revenue"));
        assert_eq!(chart_blocks[0].series[0].values, vec![100.0, 200.0]);
    }

    #[test]
    fn test_parse_docx_with_pie_chart() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rId4"/>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>"#;

        let chart_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:cat><c:strLit>
            <c:pt idx="0"><c:v>A</c:v></c:pt>
            <c:pt idx="1"><c:v>B</c:v></c:pt>
            <c:pt idx="2"><c:v>C</c:v></c:pt>
          </c:strLit></c:cat>
          <c:val><c:numLit>
            <c:pt idx="0"><c:v>30</c:v></c:pt>
            <c:pt idx="1"><c:v>50</c:v></c:pt>
            <c:pt idx="2"><c:v>20</c:v></c:pt>
          </c:numLit></c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

        let data = build_docx_with_chart(document_xml, chart_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser
            .parse(&data, &ConvertOptions::default())
            .expect("parse should succeed");

        let content = match &doc.pages[0] {
            crate::ir::Page::Flow(fp) => &fp.content,
            _ => panic!("Expected FlowPage"),
        };
        let chart_blocks: Vec<&crate::ir::Chart> = content
            .iter()
            .filter_map(|b| match b {
                Block::Chart(c) => Some(c),
                _ => None,
            })
            .collect();

        assert_eq!(chart_blocks.len(), 1);
        assert_eq!(chart_blocks[0].chart_type, crate::ir::ChartType::Pie);
        assert!(chart_blocks[0].title.is_none());
        assert_eq!(chart_blocks[0].categories, vec!["A", "B", "C"]);
        assert_eq!(chart_blocks[0].series[0].values, vec![30.0, 50.0, 20.0]);
    }

    // ── Metadata extraction tests ──────────────────────────────────────

    /// Build a minimal DOCX with docProps/core.xml containing metadata.
    fn build_docx_with_metadata(core_xml: &str) -> Vec<u8> {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        let options = zip::write::FileOptions::default();

        zip.start_file("[Content_Types].xml", options).unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
        )
        .unwrap();

        zip.start_file("_rels/.rels", options).unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
        )
        .unwrap();

        zip.start_file("word/_rels/document.xml.rels", options)
            .unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        )
        .unwrap();

        zip.start_file("word/document.xml", options).unwrap();
        std::io::Write::write_all(
            &mut zip,
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#,
        )
        .unwrap();

        zip.start_file("docProps/core.xml", options).unwrap();
        std::io::Write::write_all(&mut zip, core_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn test_parse_docx_extracts_metadata() {
        let core_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>My DOCX Title</dc:title>
  <dc:creator>DOCX Author</dc:creator>
  <dc:subject>DOCX Subject</dc:subject>
  <dc:description>DOCX description text</dc:description>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-03-15T08:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-04-20T12:30:00Z</dcterms:modified>
</cp:coreProperties>"#;

        let data = build_docx_with_metadata(core_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(doc.metadata.title.as_deref(), Some("My DOCX Title"));
        assert_eq!(doc.metadata.author.as_deref(), Some("DOCX Author"));
        assert_eq!(doc.metadata.subject.as_deref(), Some("DOCX Subject"));
        assert_eq!(
            doc.metadata.description.as_deref(),
            Some("DOCX description text")
        );
        assert_eq!(
            doc.metadata.created.as_deref(),
            Some("2024-03-15T08:00:00Z")
        );
        assert_eq!(
            doc.metadata.modified.as_deref(),
            Some("2024-04-20T12:30:00Z")
        );
    }

    #[test]
    fn test_parse_docx_without_metadata_no_crash() {
        // A minimal DOCX without docProps/core.xml → defaults
        let data = build_docx_with_math(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>No metadata</w:t></w:r></w:p>
        <w:sectPr/>
    </w:body>
</w:document>"#,
        );
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        // Should not crash; fields are None
        assert!(doc.metadata.title.is_none());
        assert!(doc.metadata.author.is_none());
    }

    // --- Heading level IR tests (US-096) ---

    #[test]
    fn test_heading1_sets_heading_level_in_ir() {
        let h1_style = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph)
            .name("Heading 1")
            .outline_lvl(0);

        let data = build_docx_bytes_with_styles(
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Title"))
                    .style("Heading1"),
            ],
            vec![h1_style],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(
            para.style.heading_level,
            Some(1),
            "Heading 1 (outline_lvl 0) should set heading_level = 1"
        );
    }

    #[test]
    fn test_heading2_sets_heading_level_in_ir() {
        let h2_style = docx_rs::Style::new("Heading2", docx_rs::StyleType::Paragraph)
            .name("Heading 2")
            .outline_lvl(1);

        let data = build_docx_bytes_with_styles(
            vec![
                docx_rs::Paragraph::new()
                    .add_run(docx_rs::Run::new().add_text("Subtitle"))
                    .style("Heading2"),
            ],
            vec![h2_style],
        );

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(
            para.style.heading_level,
            Some(2),
            "Heading 2 (outline_lvl 1) should set heading_level = 2"
        );
    }

    #[test]
    fn test_normal_paragraph_no_heading_level() {
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Normal text")),
        ]);

        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();
        let para = first_paragraph(&doc);
        assert_eq!(
            para.style.heading_level, None,
            "Normal paragraph should not have heading_level"
        );
    }

    // --- US-103: Multi-column section layout tests ---

    /// Helper: build a DOCX from raw document.xml (reuses build_docx_with_math pattern)
    fn build_docx_with_columns(document_xml: &str) -> Vec<u8> {
        build_docx_with_math(document_xml)
    }

    #[test]
    fn test_parse_docx_two_column_equal() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Column content</w:t></w:r></w:p>
        <w:sectPr>
            <w:cols w:num="2" w:space="720"/>
        </w:sectPr>
    </w:body>
</w:document>"#;
        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };
        let cols = flow.columns.as_ref().expect("Should have column layout");
        assert_eq!(cols.num_columns, 2);
        // 720 twips = 36pt
        assert!(
            (cols.spacing - 36.0).abs() < 0.1,
            "spacing: {}",
            cols.spacing
        );
        assert!(
            cols.column_widths.is_none(),
            "Equal columns should not have per-column widths"
        );
    }

    #[test]
    fn test_parse_docx_section_specific_column_layouts() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Section one intro</w:t></w:r></w:p>
        <w:p>
            <w:pPr>
                <w:sectPr>
                    <w:cols w:num="2" w:space="720"/>
                </w:sectPr>
            </w:pPr>
            <w:r><w:t>Section one end</w:t></w:r>
        </w:p>
        <w:p><w:r><w:t>Section two content</w:t></w:r></w:p>
        <w:sectPr>
            <w:cols w:num="1" w:space="720"/>
        </w:sectPr>
    </w:body>
</w:document>"#;
        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        assert_eq!(doc.pages.len(), 2, "Expected one FlowPage per section");

        let first = match &doc.pages[0] {
            Page::Flow(flow) => flow,
            _ => panic!("Expected FlowPage"),
        };
        let second = match &doc.pages[1] {
            Page::Flow(flow) => flow,
            _ => panic!("Expected FlowPage"),
        };

        assert_eq!(
            first.columns.as_ref().map(|layout| layout.num_columns),
            Some(2),
            "First section should keep the two-column layout"
        );
        assert!(
            second.columns.is_none(),
            "Final single-column section should not expose a column layout"
        );
    }

    #[test]
    fn test_parse_docx_three_column_equal() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Content</w:t></w:r></w:p>
        <w:sectPr>
            <w:cols w:num="3" w:space="360"/>
        </w:sectPr>
    </w:body>
</w:document>"#;
        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };
        let cols = flow.columns.as_ref().expect("Should have column layout");
        assert_eq!(cols.num_columns, 3);
        // 360 twips = 18pt
        assert!((cols.spacing - 18.0).abs() < 0.1);
    }

    #[test]
    fn test_parse_docx_unequal_columns() {
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Content</w:t></w:r></w:p>
        <w:sectPr>
            <w:cols w:num="2" w:space="720" w:equalWidth="0">
                <w:col w:w="6000" w:space="720"/>
                <w:col w:w="3000"/>
            </w:cols>
        </w:sectPr>
    </w:body>
</w:document>"#;
        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };
        let cols = flow.columns.as_ref().expect("Should have column layout");
        assert_eq!(cols.num_columns, 2);
        let widths = cols
            .column_widths
            .as_ref()
            .expect("Should have per-column widths");
        assert_eq!(widths.len(), 2);
        // 6000 twips = 300pt, 3000 twips = 150pt
        assert!((widths[0] - 300.0).abs() < 0.1, "width[0]: {}", widths[0]);
        assert!((widths[1] - 150.0).abs() < 0.1, "width[1]: {}", widths[1]);
    }

    #[test]
    fn test_parse_docx_no_columns() {
        // Default document without w:cols should have no column layout
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Normal")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };
        assert!(
            flow.columns.is_none(),
            "Normal doc should not have column layout"
        );
    }

    #[test]
    fn test_parse_docx_column_break() {
        // Test that w:br with type="column" produces Block::ColumnBreak
        let data = build_docx_bytes(vec![
            docx_rs::Paragraph::new()
                .add_run(docx_rs::Run::new().add_text("Before"))
                .add_run(docx_rs::Run::new().add_break(docx_rs::BreakType::Column))
                .add_run(docx_rs::Run::new().add_text("After")),
        ]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };

        // Should have a ColumnBreak block
        let has_col_break = flow.content.iter().any(|b| matches!(b, Block::ColumnBreak));
        assert!(
            has_col_break,
            "Should have a ColumnBreak block. Blocks: {:?}",
            flow.content
                .iter()
                .map(std::mem::discriminant)
                .collect::<Vec<_>>()
        );
    }

    #[test]
    fn test_parse_docx_single_column_no_layout() {
        // w:cols with num="1" should not produce column layout
        let document_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p><w:r><w:t>Content</w:t></w:r></w:p>
        <w:sectPr>
            <w:cols w:num="1" w:space="720"/>
        </w:sectPr>
    </w:body>
</w:document>"#;
        let data = build_docx_with_columns(document_xml);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };
        assert!(
            flow.columns.is_none(),
            "Single column should not produce column layout"
        );
    }

    #[test]
    fn test_extract_tab_stops_preserves_explicit_clear_override() {
        let tabs = vec![
            docx_rs::Tab::new()
                .val(docx_rs::TabValueType::Clear)
                .pos(1440),
        ];

        let tab_stops = extract_tab_stops(&tabs);

        assert_eq!(
            tab_stops,
            Some(vec![]),
            "A paragraph-level clear tab must remain an explicit empty override"
        );
    }

    #[test]
    fn test_merge_paragraph_style_preserves_inherited_tabs_not_overridden() {
        let explicit_prop = docx_rs::ParagraphProperty::new().add_tab(
            docx_rs::Tab::new()
                .val(docx_rs::TabValueType::Left)
                .pos(2160),
        );
        let explicit = extract_paragraph_style(&explicit_prop);
        let explicit_tab_overrides = extract_tab_stop_overrides(&explicit_prop.tabs);
        let style = ResolvedStyle {
            text: TextStyle::default(),
            paragraph: ParagraphStyle {
                tab_stops: Some(vec![
                    TabStop {
                        position: 72.0,
                        alignment: TabAlignment::Left,
                        leader: TabLeader::None,
                    },
                    TabStop {
                        position: 144.0,
                        alignment: TabAlignment::Right,
                        leader: TabLeader::Dot,
                    },
                ]),
                ..ParagraphStyle::default()
            },
            paragraph_tab_overrides: None,
            heading_level: None,
        };

        let merged =
            merge_paragraph_style(&explicit, explicit_tab_overrides.as_deref(), Some(&style));

        assert_eq!(
            merged.tab_stops,
            Some(vec![
                TabStop {
                    position: 72.0,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                },
                TabStop {
                    position: 108.0,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                },
                TabStop {
                    position: 144.0,
                    alignment: TabAlignment::Right,
                    leader: TabLeader::Dot,
                },
            ]),
            "Paragraph-level tabs should extend inherited style tabs instead of replacing them"
        );
    }

    #[test]
    fn test_merge_paragraph_style_clears_only_targeted_inherited_tab_stop() {
        let explicit_prop = docx_rs::ParagraphProperty::new()
            .add_tab(
                docx_rs::Tab::new()
                    .val(docx_rs::TabValueType::Clear)
                    .pos(2880),
            )
            .add_tab(
                docx_rs::Tab::new()
                    .val(docx_rs::TabValueType::Left)
                    .pos(2160),
            );
        let explicit = extract_paragraph_style(&explicit_prop);
        let explicit_tab_overrides = extract_tab_stop_overrides(&explicit_prop.tabs);
        let style = ResolvedStyle {
            text: TextStyle::default(),
            paragraph: ParagraphStyle {
                tab_stops: Some(vec![
                    TabStop {
                        position: 72.0,
                        alignment: TabAlignment::Left,
                        leader: TabLeader::None,
                    },
                    TabStop {
                        position: 144.0,
                        alignment: TabAlignment::Right,
                        leader: TabLeader::Dot,
                    },
                ]),
                ..ParagraphStyle::default()
            },
            paragraph_tab_overrides: None,
            heading_level: None,
        };

        let merged =
            merge_paragraph_style(&explicit, explicit_tab_overrides.as_deref(), Some(&style));

        assert_eq!(
            merged.tab_stops,
            Some(vec![
                TabStop {
                    position: 72.0,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                },
                TabStop {
                    position: 108.0,
                    alignment: TabAlignment::Left,
                    leader: TabLeader::None,
                },
            ]),
            "A clear tab should remove only the matching inherited stop, not the whole inherited list"
        );
    }

    #[test]
    fn test_merge_paragraph_style_allows_clearing_inherited_tab_stops() {
        let inherited = TabStop {
            position: 72.0,
            alignment: TabAlignment::Left,
            leader: TabLeader::None,
        };
        let explicit = ParagraphStyle {
            tab_stops: Some(vec![]),
            ..ParagraphStyle::default()
        };
        let style = ResolvedStyle {
            text: TextStyle::default(),
            paragraph: ParagraphStyle {
                tab_stops: Some(vec![inherited]),
                ..ParagraphStyle::default()
            },
            paragraph_tab_overrides: None,
            heading_level: None,
        };

        let merged = merge_paragraph_style(&explicit, None, Some(&style));

        assert_eq!(
            merged.tab_stops,
            Some(vec![]),
            "Explicit paragraph tab clearing must override inherited style tab stops"
        );
    }

    // ── BiDi / RTL tests ──────────────────────────────────────────────

    /// Helper: create a bidi paragraph with the given text.
    fn make_bidi_paragraph(text: &str) -> docx_rs::Paragraph {
        let mut para = docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text(text));
        para.property = docx_rs::ParagraphProperty::new().bidi(true);
        para
    }

    #[test]
    fn test_parse_docx_bidi_paragraph() {
        // Build a DOCX with a bidi paragraph containing Arabic text
        let para = make_bidi_paragraph("مرحبا بالعالم");
        let data = build_docx_bytes(vec![para]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };
        let para_block = flow.content.iter().find_map(|b| match b {
            Block::Paragraph(p) => Some(p),
            _ => None,
        });
        let p = para_block.expect("Should have a paragraph");
        assert_eq!(
            p.style.direction,
            Some(TextDirection::Rtl),
            "bidi paragraph should have RTL direction"
        );
    }

    #[test]
    fn test_parse_docx_no_bidi_paragraph() {
        // Normal LTR paragraph should have direction: None
        let para = docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello World"));
        let data = build_docx_bytes(vec![para]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };
        let para_block = flow.content.iter().find_map(|b| match b {
            Block::Paragraph(p) => Some(p),
            _ => None,
        });
        let p = para_block.expect("Should have a paragraph");
        assert!(
            p.style.direction.is_none(),
            "Non-bidi paragraph should have no direction"
        );
    }

    #[test]
    fn test_parse_docx_mixed_bidi_paragraphs() {
        // Mixed: first paragraph is RTL Arabic, second is LTR English
        let para_rtl = make_bidi_paragraph("مرحبا 123");
        let para_ltr =
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello World"));
        let data = build_docx_bytes(vec![para_rtl, para_ltr]);
        let parser = DocxParser;
        let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

        let flow = match &doc.pages[0] {
            Page::Flow(f) => f,
            _ => panic!("Expected FlowPage"),
        };
        let paras: Vec<&Paragraph> = flow
            .content
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => Some(p),
                _ => None,
            })
            .collect();
        assert!(paras.len() >= 2, "Should have at least 2 paragraphs");
        assert_eq!(
            paras[0].style.direction,
            Some(TextDirection::Rtl),
            "First paragraph (Arabic) should be RTL"
        );
        assert!(
            paras[1].style.direction.is_none(),
            "Second paragraph (English) should have no direction"
        );
    }

    #[test]
    fn test_resolve_highlight_color_named_colors() {
        assert_eq!(
            resolve_highlight_color("yellow"),
            Some(Color::new(255, 255, 0))
        );
        assert_eq!(
            resolve_highlight_color("green"),
            Some(Color::new(0, 255, 0))
        );
        assert_eq!(
            resolve_highlight_color("cyan"),
            Some(Color::new(0, 255, 255))
        );
        assert_eq!(resolve_highlight_color("red"), Some(Color::new(255, 0, 0)));
        assert_eq!(
            resolve_highlight_color("darkBlue"),
            Some(Color::new(0, 0, 128))
        );
        assert_eq!(resolve_highlight_color("black"), Some(Color::new(0, 0, 0)));
        assert_eq!(
            resolve_highlight_color("white"),
            Some(Color::new(255, 255, 255))
        );
        assert_eq!(resolve_highlight_color("none"), None);
        assert_eq!(resolve_highlight_color("unknown"), None);
    }

    #[test]
    fn test_highlight_parsing_from_docx() {
        let para = docx_rs::Paragraph::new().add_run(
            docx_rs::Run::new()
                .add_text("Highlighted")
                .highlight("yellow"),
        );
        let data: Vec<u8> = build_docx_bytes(vec![para]);
        let (doc, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
        let pages: Vec<&FlowPage> = doc
            .pages
            .iter()
            .filter_map(|p| match p {
                Page::Flow(fp) => Some(fp),
                _ => None,
            })
            .collect();
        let runs: Vec<&Run> = pages
            .iter()
            .flat_map(|p| &p.content)
            .filter_map(|b| match b {
                Block::Paragraph(p) => Some(&p.runs),
                _ => None,
            })
            .flatten()
            .collect();
        let highlighted: Vec<&&Run> = runs
            .iter()
            .filter(|r| r.style.highlight.is_some())
            .collect();
        assert!(
            !highlighted.is_empty(),
            "Should have at least one run with highlight color"
        );
        assert_eq!(
            highlighted[0].style.highlight,
            Some(Color::new(255, 255, 0)),
            "Yellow highlight should map to (255, 255, 0)"
        );
    }
}
