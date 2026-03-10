use std::cell::Cell;
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};

/// Maximum nesting depth for tables-within-tables.  Deeper nesting is silently
/// truncated to prevent stack overflow on pathological documents.
const MAX_TABLE_DEPTH: usize = 64;
use crate::ir::{
    Alignment, Block, BorderLineStyle, BorderSide, CellBorder, CellVerticalAlign, Chart, Color,
    ColumnLayout, Document, FloatingImage, FloatingTextBox, ImageData, ImageFormat, Insets,
    LineSpacing, MathEquation, Page, Paragraph, ParagraphStyle, Run, StyleSheet, TabAlignment,
    TabLeader, TabStop, Table, TableCell, TableRow, TextDirection, TextStyle, VerticalTextAlign,
    WrapMode,
};
use crate::parser::Parser;

use self::lists::{
    NumberingMap, TaggedElement, build_numbering_map, extract_num_info, group_into_lists,
};
#[cfg(test)]
use self::sections::extract_page_size;
use self::sections::{
    HeaderFooterAssets, build_flow_page_from_section, build_header_footer_assets,
};
use self::styles::{
    DOC_DEFAULT_STYLE_ID, ResolvedStyle, StyleMap, TabStopOverride, apply_tab_stop_overrides,
    build_style_map, get_paragraph_style_id, merge_paragraph_style, merge_text_style,
};

#[path = "docx_lists.rs"]
mod lists;
#[path = "docx_sections.rs"]
mod sections;
#[path = "docx_styles.rs"]
mod styles;

/// Parser for DOCX (Office Open XML Word) documents.
pub struct DocxParser;

/// Map from relationship ID → PNG image bytes.
type ImageMap = HashMap<String, Vec<u8>>;

/// Map from relationship ID → hyperlink URL.
type HyperlinkMap = HashMap<String, String>;

/// Build a lookup map from the DOCX's hyperlinks (reader-populated field).
/// The reader stores hyperlinks as `(rid, url, type)` in `docx.hyperlinks`.
fn build_hyperlink_map(docx: &docx_rs::Docx) -> HyperlinkMap {
    docx.hyperlinks
        .iter()
        .map(|(rid, url, _type)| (rid.clone(), url.clone()))
        .collect()
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

fn extract_table_width_points(prop_json: Option<&serde_json::Value>) -> Option<f64> {
    let width_json = prop_json.and_then(|json| json.get("width"))?;
    let width_type = width_json
        .get("widthType")
        .and_then(|value| value.as_str())?;
    if width_type != "dxa" {
        return None;
    }
    width_json
        .get("width")
        .and_then(|value| value.as_f64())
        .map(|width| width / 20.0)
}

fn is_auto_table_width(prop_json: Option<&serde_json::Value>) -> bool {
    prop_json
        .and_then(|json| json.get("width"))
        .and_then(|width| width.get("widthType"))
        .and_then(|value| value.as_str())
        == Some("auto")
}

fn has_placeholder_autofit_grid(
    table: &docx_rs::Table,
    prop_json: Option<&serde_json::Value>,
) -> bool {
    if table.grid.is_empty() || !is_auto_table_width(prop_json) {
        return false;
    }

    // Some generators emit gridCol=100 twips (5pt) placeholders for auto-fit
    // tables. Treat these as non-authoritative to avoid collapsed columns.
    const PLACEHOLDER_GRID_MAX_TWIPS: usize = 200;
    table
        .grid
        .iter()
        .all(|column_width_twips| *column_width_twips <= PLACEHOLDER_GRID_MAX_TWIPS)
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

    let inferred_column_widths: Option<Vec<f64>> = derive_column_widths_from_cells(&raw_rows);
    let mut grid_widths_points: Vec<f64> = table
        .grid
        .iter()
        .map(|&width| width as f64 / 20.0)
        .collect();

    let column_widths: Vec<f64> = if table.grid.is_empty()
        || has_placeholder_autofit_grid(table, table_prop_json.as_ref())
    {
        inferred_column_widths.unwrap_or_default()
    } else if let Some(table_width_points) = extract_table_width_points(table_prop_json.as_ref()) {
        let grid_total_points: f64 = grid_widths_points.iter().sum();
        if grid_total_points > 0.0 {
            let scale = table_width_points / grid_total_points;
            // Keep authored grid ratios but normalize to table width to avoid
            // collapsed placeholder grids (e.g. all gridCol=100 twips).
            if (scale - 1.0).abs() > 0.05 {
                for width in &mut grid_widths_points {
                    *width *= scale;
                }
            }
            grid_widths_points
        } else if let Some(widths) = inferred_column_widths {
            widths
        } else {
            let column_count = table.grid.len().max(1);
            vec![table_width_points / column_count as f64; column_count]
        }
    } else {
        grid_widths_points
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

/// Extract document-level default paragraph style from styles.xml docDefaults.
fn extract_doc_default_paragraph_style(styles: &docx_rs::Styles) -> ParagraphStyle {
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
        .and_then(|value| value.get("val"))
        .and_then(serde_json::Value::as_str)
        .and_then(|value| match value {
            "center" => Some(Alignment::Center),
            "right" | "end" => Some(Alignment::Right),
            "left" | "start" => Some(Alignment::Left),
            "both" | "justified" => Some(Alignment::Justify),
            _ => None,
        });

    let indent = paragraph_property.get("indent");
    let indent_left = indent
        .and_then(|value| value.get("start"))
        .and_then(serde_json::Value::as_f64)
        .map(|value| value / 20.0);
    let indent_right = indent
        .and_then(|value| value.get("end"))
        .and_then(serde_json::Value::as_f64)
        .map(|value| value / 20.0);
    let indent_first_line = indent
        .and_then(|value| value.get("specialIndent"))
        .and_then(|value| {
            value
                .get("firstLine")
                .and_then(serde_json::Value::as_f64)
                .map(|twips| twips / 20.0)
                .or_else(|| {
                    value
                        .get("hanging")
                        .and_then(serde_json::Value::as_f64)
                        .map(|twips| -(twips / 20.0))
                })
        });

    let (line_spacing, space_before, space_after) =
        extract_line_spacing_from_json(paragraph_property.get("lineSpacing"));

    ParagraphStyle {
        alignment,
        indent_left,
        indent_right,
        indent_first_line,
        line_spacing,
        space_before,
        space_after,
        heading_level: None,
        direction: paragraph_property
            .get("bidirectional")
            .and_then(json_bool_or_val)
            .and_then(|is_rtl| is_rtl.then_some(TextDirection::Rtl)),
        tab_stops: None,
    }
}

fn extract_line_spacing_from_json(
    spacing: Option<&serde_json::Value>,
) -> (Option<LineSpacing>, Option<f64>, Option<f64>) {
    let Some(spacing) = spacing else {
        return (None, None, None);
    };

    let space_before = spacing
        .get("before")
        .and_then(serde_json::Value::as_f64)
        .map(|value| value / 20.0);
    let space_after = spacing
        .get("after")
        .and_then(serde_json::Value::as_f64)
        .map(|value| value / 20.0);

    let line_spacing = spacing
        .get("line")
        .and_then(serde_json::Value::as_f64)
        .map(|line| {
            match spacing
                .get("lineRule")
                .and_then(serde_json::Value::as_str)
                .unwrap_or("auto")
            {
                "exact" | "atLeast" => LineSpacing::Exact(line / 20.0),
                _ => LineSpacing::Proportional(line / 240.0),
            }
        });

    (line_spacing, space_before, space_after)
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
#[path = "docx_tests.rs"]
mod tests;
