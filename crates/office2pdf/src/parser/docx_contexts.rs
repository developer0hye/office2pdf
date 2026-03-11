use std::cell::Cell;
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};

use quick_xml::events::BytesStart;

use crate::ir::{
    Block, BorderLineStyle, Chart, Color, ColumnLayout, Insets, MathEquation, Paragraph,
    ParagraphBorder, ParagraphBorderSide, ParagraphContainerStyle, ParagraphStyle, Run, TextStyle,
    WrapMode,
};
use crate::parser::{chart, omml};

use super::{emu_to_pt, extract_run_text, parse_hex_color};

// ── Footnote / Endnote support ──────────────────────────────────────────

#[derive(Debug, Clone, Copy)]
enum NoteKind {
    Footnote,
    Endnote,
}

/// Context for resolving footnote/endnote references during parsing.
/// The `cursor` is advanced each time a note reference run is encountered.
pub(super) struct NoteContext {
    footnote_content: HashMap<usize, String>,
    endnote_content: HashMap<usize, String>,
    note_refs: Vec<(NoteKind, usize)>,
    cursor: Cell<usize>,
    note_style_ids: HashSet<String>,
}

impl NoteContext {
    pub(super) fn empty() -> Self {
        let note_style_ids: HashSet<String> = ["FootnoteReference", "EndnoteReference"]
            .iter()
            .map(|style_id| (*style_id).to_string())
            .collect();
        Self {
            footnote_content: HashMap::new(),
            endnote_content: HashMap::new(),
            note_refs: Vec::new(),
            cursor: Cell::new(0),
            note_style_ids,
        }
    }

    pub(super) fn consume_next(&self) -> Option<String> {
        let index = self.cursor.get();
        if index >= self.note_refs.len() {
            return None;
        }
        let (kind, id) = self.note_refs[index];
        self.cursor.set(index + 1);
        match kind {
            NoteKind::Footnote => self.footnote_content.get(&id).cloned(),
            NoteKind::Endnote => self.endnote_content.get(&id).cloned(),
        }
    }

    pub(super) fn populate_style_ids(&mut self, styles: &docx_rs::Styles) {
        for style in &styles.styles {
            if let Ok(name_value) = serde_json::to_value(&style.name)
                && let Some(name_str) = name_value.as_str()
            {
                let lower = name_str.to_lowercase();
                if lower == "footnote reference" || lower == "endnote reference" {
                    self.note_style_ids.insert(style.style_id.clone());
                }
            }
        }
    }
}

struct AnchorWrapInfo {
    wrap_mode: WrapMode,
    behind_doc: bool,
}

pub(super) struct WrapContext {
    wraps: Vec<AnchorWrapInfo>,
    cursor: Cell<usize>,
}

impl WrapContext {
    pub(super) fn empty() -> Self {
        Self {
            wraps: Vec::new(),
            cursor: Cell::new(0),
        }
    }

    pub(super) fn consume_next(&self) -> WrapMode {
        let index = self.cursor.get();
        if index >= self.wraps.len() {
            return WrapMode::None;
        }
        let info = &self.wraps[index];
        self.cursor.set(index + 1);
        if info.behind_doc {
            WrapMode::Behind
        } else {
            info.wrap_mode
        }
    }
}

pub(super) fn build_wrap_context_from_xml(doc_xml: Option<&str>) -> WrapContext {
    let wraps = doc_xml.map(scan_anchor_wrap_types).unwrap_or_default();
    WrapContext {
        wraps,
        cursor: Cell::new(0),
    }
}

#[derive(Debug, Clone, Copy, Default)]
pub(super) struct DrawingTextBoxInfo {
    pub(super) width_pt: Option<f64>,
    pub(super) height_pt: Option<f64>,
}

pub(super) struct DrawingTextBoxContext {
    text_boxes: Vec<DrawingTextBoxInfo>,
    cursor: Cell<usize>,
}

impl DrawingTextBoxContext {
    pub(super) fn from_xml(xml: Option<&str>) -> Self {
        Self {
            text_boxes: xml.map(scan_drawing_text_boxes).unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    pub(super) fn consume_next(&self) -> DrawingTextBoxInfo {
        let index = self.cursor.get();
        self.cursor.set(index + 1);
        self.text_boxes.get(index).copied().unwrap_or_default()
    }
}

#[derive(Debug, Clone, Copy, Default)]
pub(super) struct TableHeaderInfo {
    pub(super) repeat_rows: usize,
}

pub(super) struct TableHeaderContext {
    headers: Vec<TableHeaderInfo>,
    cursor: Cell<usize>,
}

impl TableHeaderContext {
    pub(super) fn from_xml(xml: Option<&str>) -> Self {
        Self {
            headers: xml.map(scan_table_headers).unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    pub(super) fn consume_next(&self) -> TableHeaderInfo {
        let index = self.cursor.get();
        self.cursor.set(index + 1);
        self.headers.get(index).copied().unwrap_or_default()
    }
}

struct TableHeaderScanState {
    table_index: usize,
    repeat_rows: usize,
    in_row: bool,
    current_row_is_header: bool,
    saw_body_row: bool,
}

#[cfg(test)]
pub(super) fn scan_table_headers(xml: &str) -> Vec<TableHeaderInfo> {
    scan_table_headers_impl(xml)
}

#[cfg(not(test))]
pub(super) fn scan_table_headers(xml: &str) -> Vec<TableHeaderInfo> {
    scan_table_headers_impl(xml)
}

fn scan_table_headers_impl(xml: &str) -> Vec<TableHeaderInfo> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut headers: Vec<TableHeaderInfo> = Vec::new();
    let mut stack: Vec<TableHeaderScanState> = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(quick_xml::events::Event::Start(ref element)) => match element.local_name().as_ref()
            {
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
                        && on_off_element_is_enabled(element)
                    {
                        state.current_row_is_header = true;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref element)) => match element.local_name().as_ref()
            {
                b"tbl" => headers.push(TableHeaderInfo::default()),
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
                        && on_off_element_is_enabled(element)
                    {
                        state.current_row_is_header = true;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
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
        buffer.clear();
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

fn on_off_element_is_enabled(element: &quick_xml::events::BytesStart<'_>) -> bool {
    for attribute in element.attributes().flatten() {
        if attribute.key.local_name().as_ref() != b"val" {
            continue;
        }

        let value = attribute.value.as_ref();
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
    let mut buffer: Vec<u8> = Vec::new();
    let mut result: Vec<DrawingTextBoxInfo> = Vec::new();
    let mut in_body: bool = false;
    let mut drawing_depth: usize = 0;
    let mut current_info: DrawingTextBoxInfo = DrawingTextBoxInfo::default();
    let mut saw_text_box: bool = false;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(quick_xml::events::Event::Start(ref element)) => match element.local_name().as_ref()
            {
                b"body" => in_body = true,
                b"drawing" if in_body => {
                    if drawing_depth == 0 {
                        current_info = DrawingTextBoxInfo::default();
                        saw_text_box = false;
                    }
                    drawing_depth += 1;
                }
                b"extent" if drawing_depth > 0 => {
                    update_drawing_text_box_extent(&mut current_info, element);
                }
                b"txbx" if drawing_depth > 0 => saw_text_box = true,
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref element)) => match element.local_name().as_ref()
            {
                b"extent" if drawing_depth > 0 => {
                    update_drawing_text_box_extent(&mut current_info, element);
                }
                b"txbx" if drawing_depth > 0 => saw_text_box = true,
                _ => {}
            },
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
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
        buffer.clear();
    }

    result
}

fn update_drawing_text_box_extent(
    info: &mut DrawingTextBoxInfo,
    element: &quick_xml::events::BytesStart<'_>,
) {
    if info.width_pt.is_some() && info.height_pt.is_some() {
        return;
    }

    let mut width_emu: Option<u32> = None;
    let mut height_emu: Option<u32> = None;

    for attribute in element.attributes().flatten() {
        match attribute.key.local_name().as_ref() {
            b"cx" => {
                width_emu = std::str::from_utf8(attribute.value.as_ref())
                    .ok()
                    .and_then(|value| value.parse::<u32>().ok());
            }
            b"cy" => {
                height_emu = std::str::from_utf8(attribute.value.as_ref())
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
pub(super) struct VmlTextBoxInfo {
    pub(super) paragraphs: Vec<String>,
    pub(super) wrap_mode: Option<WrapMode>,
}

impl VmlTextBoxInfo {
    pub(super) fn into_blocks(self) -> Vec<Block> {
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

pub(super) struct VmlTextBoxContext {
    text_boxes: Vec<VmlTextBoxInfo>,
    cursor: Cell<usize>,
}

impl VmlTextBoxContext {
    pub(super) fn from_xml(xml: Option<&str>) -> Self {
        Self {
            text_boxes: xml.map(scan_vml_text_boxes).unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    pub(super) fn consume_next(&self) -> VmlTextBoxInfo {
        let index: usize = self.cursor.get();
        self.cursor.set(index + 1);
        self.text_boxes.get(index).cloned().unwrap_or_default()
    }
}

fn scan_vml_text_boxes(xml: &str) -> Vec<VmlTextBoxInfo> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut result: Vec<VmlTextBoxInfo> = Vec::new();
    let mut in_body: bool = false;
    let mut pict_depth: usize = 0;
    let mut shape_depth: usize = 0;
    let mut in_text_box_content: bool = false;
    let mut in_paragraph: bool = false;
    let mut current_picture_shapes: Vec<VmlTextBoxInfo> = Vec::new();
    let mut current_picture_wrap: Option<WrapMode> = None;
    let mut current_shape_paragraphs: Vec<String> = Vec::new();
    let mut current_paragraph_text: String = String::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(quick_xml::events::Event::Start(ref element)) => match element.local_name().as_ref()
            {
                b"body" => in_body = true,
                b"pict" if in_body => {
                    if pict_depth == 0 {
                        current_picture_shapes.clear();
                        current_picture_wrap = None;
                    }
                    pict_depth += 1;
                }
                b"shape" if pict_depth > 0 => {
                    if shape_depth == 0 {
                        current_shape_paragraphs.clear();
                    }
                    shape_depth += 1;
                }
                b"txbxContent" if shape_depth > 0 => in_text_box_content = true,
                b"p" if in_text_box_content => {
                    in_paragraph = true;
                    current_paragraph_text.clear();
                }
                b"wrap" if pict_depth > 0 => {
                    current_picture_wrap = extract_vml_wrap_mode_from_element(element);
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref element)) => match element.local_name().as_ref()
            {
                b"tab" if in_paragraph => current_paragraph_text.push('\t'),
                b"br" if in_paragraph => current_paragraph_text.push('\n'),
                b"wrap" if pict_depth > 0 => {
                    current_picture_wrap = extract_vml_wrap_mode_from_element(element);
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Text(ref element)) => {
                if in_paragraph && let Ok(text) = element.xml_content() {
                    current_paragraph_text.push_str(&text);
                }
            }
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
                b"body" => in_body = false,
                b"p" if in_paragraph => {
                    current_shape_paragraphs.push(std::mem::take(&mut current_paragraph_text));
                    in_paragraph = false;
                }
                b"txbxContent" if in_text_box_content => in_text_box_content = false,
                b"shape" if shape_depth > 0 => {
                    shape_depth -= 1;
                    if shape_depth == 0 {
                        current_picture_shapes.push(VmlTextBoxInfo {
                            paragraphs: std::mem::take(&mut current_shape_paragraphs),
                            wrap_mode: None,
                        });
                        in_text_box_content = false;
                        in_paragraph = false;
                        current_paragraph_text.clear();
                    }
                }
                b"pict" if pict_depth > 0 => {
                    pict_depth -= 1;
                    if pict_depth == 0 {
                        for mut text_box in current_picture_shapes.drain(..) {
                            text_box.wrap_mode = current_picture_wrap;
                            result.push(text_box);
                        }
                        current_picture_wrap = None;
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buffer.clear();
    }

    result
}

fn extract_vml_wrap_mode_from_element(
    element: &quick_xml::events::BytesStart<'_>,
) -> Option<WrapMode> {
    for attribute in element.attributes().flatten() {
        if attribute.key.local_name().as_ref() != b"type" {
            continue;
        }

        let value = std::str::from_utf8(attribute.value.as_ref()).ok()?;
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

fn scan_anchor_wrap_types(xml: &str) -> Vec<AnchorWrapInfo> {
    let mut results: Vec<AnchorWrapInfo> = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut in_anchor = false;
    let mut behind_doc = false;
    let mut found_wrap = false;
    let mut current_wrap = WrapMode::None;

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element))
            | Ok(quick_xml::events::Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"anchor" => {
                        in_anchor = true;
                        behind_doc = false;
                        found_wrap = false;
                        current_wrap = WrapMode::None;
                        for attribute in element.attributes().flatten() {
                            if attribute.key.local_name().as_ref() == b"behindDoc"
                                && let Ok(value) = attribute.unescape_value()
                                && (value == "1" || value == "true")
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
                        current_wrap = WrapMode::Tight;
                        found_wrap = true;
                    }
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref element)) => {
                if element.local_name().as_ref() == b"anchor" && in_anchor {
                    if !found_wrap && behind_doc {
                        current_wrap = WrapMode::None;
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

pub(super) struct BidiContext {
    bidi_indices: HashSet<usize>,
    cursor: Cell<usize>,
}

impl BidiContext {
    pub(super) fn from_xml(xml: Option<&str>) -> Self {
        let bidi_indices = xml.map(Self::scan).unwrap_or_default();
        Self {
            bidi_indices,
            cursor: Cell::new(0),
        }
    }

    pub(super) fn next_is_bidi(&self) -> bool {
        let index = self.cursor.get();
        self.cursor.set(index + 1);
        self.bidi_indices.contains(&index)
    }

    fn scan(xml: &str) -> HashSet<usize> {
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut buffer: Vec<u8> = Vec::new();
        let mut result: HashSet<usize> = HashSet::new();
        let mut paragraph_index: usize = 0;
        let mut in_paragraph_properties = false;
        let mut in_body = false;

        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(quick_xml::events::Event::Start(ref element))
                | Ok(quick_xml::events::Event::Empty(ref element)) => {
                    match element.local_name().as_ref() {
                        b"body" => in_body = true,
                        b"pPr" if in_body => in_paragraph_properties = true,
                        b"bidi" if in_paragraph_properties => {
                            result.insert(paragraph_index);
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::End(ref element)) => {
                    match element.local_name().as_ref() {
                        b"body" => in_body = false,
                        b"p" if in_body => {
                            paragraph_index += 1;
                            in_paragraph_properties = false;
                        }
                        b"pPr" => in_paragraph_properties = false,
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buffer.clear();
        }

        result
    }
}

pub(super) struct SmallCapsContext {
    flags: Vec<bool>,
    cursor: Cell<usize>,
}

impl SmallCapsContext {
    pub(super) fn from_xml(xml: Option<&str>) -> Self {
        let flags = xml.map(Self::scan).unwrap_or_default();
        Self {
            flags,
            cursor: Cell::new(0),
        }
    }

    pub(super) fn next_is_small_caps(&self) -> bool {
        let index = self.cursor.get();
        self.cursor.set(index + 1);
        self.flags.get(index).copied().unwrap_or(false)
    }

    fn scan(xml: &str) -> Vec<bool> {
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut buffer: Vec<u8> = Vec::new();
        let mut result: Vec<bool> = Vec::new();
        let mut in_body = false;
        let mut in_run = false;
        let mut in_run_properties = false;
        let mut current_has_small_caps = false;

        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(quick_xml::events::Event::Start(ref element))
                | Ok(quick_xml::events::Event::Empty(ref element)) => {
                    match element.local_name().as_ref() {
                        b"body" => in_body = true,
                        b"r" if in_body => {
                            in_run = true;
                            current_has_small_caps = false;
                        }
                        b"rPr" if in_run => in_run_properties = true,
                        b"smallCaps" if in_run_properties => {
                            let is_disabled = element.attributes().flatten().any(|attribute| {
                                attribute.key.local_name().as_ref() == b"val"
                                    && matches!(attribute.value.as_ref(), b"false" | b"0")
                            });
                            if !is_disabled {
                                current_has_small_caps = true;
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::End(ref element)) => {
                    match element.local_name().as_ref() {
                        b"body" => in_body = false,
                        b"r" if in_body => {
                            result.push(current_has_small_caps);
                            in_run = false;
                            in_run_properties = false;
                            current_has_small_caps = false;
                        }
                        b"rPr" => in_run_properties = false,
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buffer.clear();
        }

        result
    }
}

#[derive(Default)]
struct ParagraphContainerAccumulator {
    background: Option<Color>,
    top: Option<ParagraphBorderSide>,
    bottom: Option<ParagraphBorderSide>,
    left: Option<ParagraphBorderSide>,
    right: Option<ParagraphBorderSide>,
    padding_top: Option<f64>,
    padding_right: Option<f64>,
    padding_bottom: Option<f64>,
    padding_left: Option<f64>,
}

impl ParagraphContainerAccumulator {
    fn set_background(&mut self, element: &BytesStart<'_>) {
        let Some(fill) = xml_attr(element, b"fill") else {
            return;
        };
        if fill.eq_ignore_ascii_case("auto") || fill.eq_ignore_ascii_case("none") {
            return;
        }
        if let Some(color) = parse_hex_color(&fill) {
            self.background = Some(color);
        }
    }

    fn set_border_side(&mut self, side_name: &[u8], element: &BytesStart<'_>) {
        let Some((side, padding)) = parse_paragraph_border_side(element) else {
            return;
        };

        match side_name {
            b"top" => {
                self.top = Some(side);
                if let Some(padding) = padding {
                    self.padding_top = Some(padding);
                }
            }
            b"bottom" => {
                self.bottom = Some(side);
                if let Some(padding) = padding {
                    self.padding_bottom = Some(padding);
                }
            }
            b"left" => {
                self.left = Some(side);
                if let Some(padding) = padding {
                    self.padding_left = Some(padding);
                }
            }
            b"right" => {
                self.right = Some(side);
                if let Some(padding) = padding {
                    self.padding_right = Some(padding);
                }
            }
            _ => {}
        }
    }

    fn into_style(self) -> Option<ParagraphContainerStyle> {
        let border = if self.top.is_some()
            || self.bottom.is_some()
            || self.left.is_some()
            || self.right.is_some()
        {
            Some(ParagraphBorder {
                top: self.top,
                bottom: self.bottom,
                left: self.left,
                right: self.right,
            })
        } else {
            None
        };

        let padding = if self.padding_top.is_some()
            || self.padding_right.is_some()
            || self.padding_bottom.is_some()
            || self.padding_left.is_some()
        {
            Some(Insets {
                top: self.padding_top.unwrap_or(0.0),
                right: self.padding_right.unwrap_or(0.0),
                bottom: self.padding_bottom.unwrap_or(0.0),
                left: self.padding_left.unwrap_or(0.0),
            })
        } else {
            None
        };

        if self.background.is_none() && border.is_none() && padding.is_none() {
            None
        } else {
            Some(ParagraphContainerStyle {
                background: self.background,
                border,
                padding,
            })
        }
    }
}

pub(super) struct ParagraphContainerContext {
    styles: Vec<Option<ParagraphContainerStyle>>,
    cursor: Cell<usize>,
}

impl ParagraphContainerContext {
    pub(super) fn from_xml(xml: Option<&str>) -> Self {
        let styles = xml.map(Self::scan).unwrap_or_default();
        Self {
            styles,
            cursor: Cell::new(0),
        }
    }

    pub(super) fn next_style(&self) -> Option<ParagraphContainerStyle> {
        let index = self.cursor.get();
        self.cursor.set(index + 1);
        self.styles.get(index).cloned().flatten()
    }

    fn scan(xml: &str) -> Vec<Option<ParagraphContainerStyle>> {
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut buffer: Vec<u8> = Vec::new();
        let mut styles: Vec<Option<ParagraphContainerStyle>> = Vec::new();
        let mut paragraph_stack: Vec<(usize, ParagraphContainerAccumulator)> = Vec::new();
        let mut in_body = false;
        let mut paragraph_property_depth: usize = 0;
        let mut paragraph_border_depth: usize = 0;

        loop {
            match reader.read_event_into(&mut buffer) {
                Ok(quick_xml::events::Event::Start(ref element)) => {
                    match element.local_name().as_ref() {
                        b"body" => in_body = true,
                        b"p" if in_body => {
                            let index = styles.len();
                            styles.push(None);
                            paragraph_stack.push((index, ParagraphContainerAccumulator::default()));
                        }
                        b"pPr" if in_body && !paragraph_stack.is_empty() => {
                            paragraph_property_depth += 1;
                        }
                        b"pBdr" if paragraph_property_depth > 0 && !paragraph_stack.is_empty() => {
                            paragraph_border_depth += 1;
                        }
                        b"shd" if paragraph_property_depth > 0 => {
                            if let Some((_, paragraph)) = paragraph_stack.last_mut() {
                                paragraph.set_background(element);
                            }
                        }
                        b"top" | b"bottom" | b"left" | b"right" if paragraph_border_depth > 0 => {
                            if let Some((_, paragraph)) = paragraph_stack.last_mut() {
                                paragraph.set_border_side(element.local_name().as_ref(), element);
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Empty(ref element)) => {
                    match element.local_name().as_ref() {
                        b"p" if in_body => styles.push(None),
                        b"pPr" if in_body && !paragraph_stack.is_empty() => {
                            paragraph_property_depth += 1;
                            paragraph_property_depth -= 1;
                        }
                        b"pBdr" if paragraph_property_depth > 0 && !paragraph_stack.is_empty() => {
                            paragraph_border_depth += 1;
                            paragraph_border_depth -= 1;
                        }
                        b"shd" if paragraph_property_depth > 0 => {
                            if let Some((_, paragraph)) = paragraph_stack.last_mut() {
                                paragraph.set_background(element);
                            }
                        }
                        b"top" | b"bottom" | b"left" | b"right" if paragraph_border_depth > 0 => {
                            if let Some((_, paragraph)) = paragraph_stack.last_mut() {
                                paragraph.set_border_side(element.local_name().as_ref(), element);
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::End(ref element)) => {
                    match element.local_name().as_ref() {
                        b"body" => in_body = false,
                        b"pPr" => {
                            paragraph_property_depth = paragraph_property_depth.saturating_sub(1)
                        }
                        b"pBdr" => {
                            paragraph_border_depth = paragraph_border_depth.saturating_sub(1)
                        }
                        b"p" if in_body => {
                            if let Some((index, paragraph)) = paragraph_stack.pop() {
                                styles[index] = paragraph.into_style();
                            }
                        }
                        _ => {}
                    }
                }
                Ok(quick_xml::events::Event::Eof) => break,
                Err(_) => break,
                _ => {}
            }
            buffer.clear();
        }

        styles
    }
}

fn xml_attr(element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    element
        .attributes()
        .flatten()
        .find(|attribute| attribute.key.local_name().as_ref() == name)
        .and_then(|attribute| attribute.unescape_value().ok())
        .map(|value| value.into_owned())
}

fn parse_paragraph_border_side(
    element: &BytesStart<'_>,
) -> Option<(ParagraphBorderSide, Option<f64>)> {
    let border_type = xml_attr(element, b"val").unwrap_or_else(|| "single".to_string());
    if matches!(border_type.as_str(), "none" | "nil") {
        return None;
    }

    let width = xml_attr(element, b"sz")
        .and_then(|value| value.parse::<f64>().ok())
        .filter(|value| *value > 0.0)?
        / 8.0;
    let padding = xml_attr(element, b"space").and_then(|value| value.parse::<f64>().ok());
    let color = xml_attr(element, b"color")
        .filter(|value| !value.eq_ignore_ascii_case("auto"))
        .as_deref()
        .and_then(parse_hex_color)
        .unwrap_or(Color::black());
    let style = match border_type.as_str() {
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

    Some((
        ParagraphBorderSide {
            width,
            color,
            style,
        },
        padding,
    ))
}

pub(super) fn scan_column_layouts(xml: &str) -> Vec<Option<ColumnLayout>> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut layouts: Vec<Option<ColumnLayout>> = Vec::new();

    let mut in_section_properties = false;
    let mut in_columns = false;
    let mut num_columns: u32 = 1;
    let mut spacing_twips: f64 = 720.0;
    let mut equal_width = true;
    let mut column_widths: Vec<f64> = Vec::new();

    let build_layout =
        |num_columns: u32, spacing_twips: f64, equal_width: bool, column_widths: &[f64]| {
            if num_columns < 2 {
                return None;
            }

            Some(ColumnLayout {
                num_columns,
                spacing: spacing_twips / 20.0,
                column_widths: if !equal_width && !column_widths.is_empty() {
                    Some(column_widths.to_vec())
                } else {
                    None
                },
            })
        };

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element)) => match element.local_name().as_ref()
            {
                b"sectPr" => {
                    in_section_properties = true;
                    num_columns = 1;
                    spacing_twips = 720.0;
                    equal_width = true;
                    column_widths.clear();
                }
                b"cols" if in_section_properties => {
                    in_columns = true;
                    for attribute in element.attributes().flatten() {
                        let key = attribute.key.local_name();
                        if let Ok(value) = attribute.unescape_value() {
                            match key.as_ref() {
                                b"num" => {
                                    if let Ok(parsed) = value.parse::<u32>() {
                                        num_columns = parsed;
                                    }
                                }
                                b"space" => {
                                    if let Ok(parsed) = value.parse::<f64>() {
                                        spacing_twips = parsed;
                                    }
                                }
                                b"equalWidth" => equal_width = value != "0",
                                _ => {}
                            }
                        }
                    }
                }
                b"col" if in_columns => {
                    for attribute in element.attributes().flatten() {
                        if attribute.key.local_name().as_ref() == b"w"
                            && let Ok(value) = attribute.unescape_value()
                            && let Ok(parsed) = value.parse::<f64>()
                        {
                            column_widths.push(parsed / 20.0);
                        }
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref element)) => match element.local_name().as_ref()
            {
                b"sectPr" => layouts.push(build_layout(1, 720.0, true, &[])),
                b"cols" if in_section_properties => {
                    in_columns = false;
                    for attribute in element.attributes().flatten() {
                        let key = attribute.key.local_name();
                        if let Ok(value) = attribute.unescape_value() {
                            match key.as_ref() {
                                b"num" => {
                                    if let Ok(parsed) = value.parse::<u32>() {
                                        num_columns = parsed;
                                    }
                                }
                                b"space" => {
                                    if let Ok(parsed) = value.parse::<f64>() {
                                        spacing_twips = parsed;
                                    }
                                }
                                b"equalWidth" => equal_width = value != "0",
                                _ => {}
                            }
                        }
                    }
                }
                b"col" if in_columns => {
                    for attribute in element.attributes().flatten() {
                        if attribute.key.local_name().as_ref() == b"w"
                            && let Ok(value) = attribute.unescape_value()
                            && let Ok(parsed) = value.parse::<f64>()
                        {
                            column_widths.push(parsed / 20.0);
                        }
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
                b"sectPr" => {
                    layouts.push(build_layout(
                        num_columns,
                        spacing_twips,
                        equal_width,
                        &column_widths,
                    ));
                    in_section_properties = false;
                }
                b"cols" => in_columns = false,
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    layouts
}

pub(super) fn extract_column_layout_from_section_property(
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

pub(super) struct MathContext {
    equations: HashMap<usize, Vec<MathEquation>>,
}

impl MathContext {
    pub(super) fn empty() -> Self {
        Self {
            equations: HashMap::new(),
        }
    }

    pub(super) fn take(&mut self, index: usize) -> Vec<MathEquation> {
        self.equations.remove(&index).unwrap_or_default()
    }
}

pub(super) fn build_math_context_from_xml(doc_xml: Option<&str>) -> MathContext {
    let mut equations: HashMap<usize, Vec<MathEquation>> = HashMap::new();

    if let Some(xml) = doc_xml {
        let raw = omml::scan_math_equations(xml);
        for (index, content, display) in raw {
            equations
                .entry(index)
                .or_default()
                .push(MathEquation { content, display });
        }
    }

    MathContext { equations }
}

pub(super) struct ChartContext {
    charts: HashMap<usize, Vec<Chart>>,
}

impl ChartContext {
    pub(super) fn empty() -> Self {
        Self {
            charts: HashMap::new(),
        }
    }

    pub(super) fn take(&mut self, index: usize) -> Vec<Chart> {
        self.charts.remove(&index).unwrap_or_default()
    }
}

pub(super) fn build_chart_context_from_xml(
    doc_xml: Option<&str>,
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
) -> ChartContext {
    let mut charts: HashMap<usize, Vec<Chart>> = HashMap::new();

    let Some(doc_xml) = doc_xml else {
        return ChartContext { charts };
    };

    let Some(relationships_xml) = read_zip_text(archive, "word/_rels/document.xml.rels") else {
        return ChartContext { charts };
    };

    let chart_references = chart::scan_chart_references(doc_xml);
    let chart_relationships = chart::scan_chart_rels(&relationships_xml);

    for (body_index, relationship_id) in chart_references {
        if let Some(chart_path) = chart_relationships.get(&relationship_id)
            && let Some(chart_xml) = read_zip_text(archive, chart_path)
            && let Some(chart) = chart::parse_chart_xml(&chart_xml)
        {
            charts.entry(body_index).or_default().push(chart);
        }
    }

    ChartContext { charts }
}

pub(super) fn build_note_context_from_xml(
    doc_xml: Option<&str>,
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
) -> NoteContext {
    let mut note_context = NoteContext::empty();

    if let Some(xml) = read_zip_text(archive, "word/footnotes.xml") {
        note_context.footnote_content = parse_notes_xml(&xml);
    }
    if let Some(xml) = read_zip_text(archive, "word/endnotes.xml") {
        note_context.endnote_content = parse_notes_xml(&xml);
    }
    note_context.note_refs = doc_xml.map(scan_note_refs).unwrap_or_default();

    note_context
}

pub(super) fn read_zip_text(
    archive: &mut zip::ZipArchive<impl Read + Seek>,
    name: &str,
) -> Option<String> {
    let mut file = archive.by_name(name).ok()?;
    let mut contents = String::new();
    file.read_to_string(&mut contents).ok()?;
    Some(contents)
}

fn parse_notes_xml(xml: &str) -> HashMap<usize, String> {
    let mut map: HashMap<usize, String> = HashMap::new();
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut current_id: Option<usize> = None;
    let mut current_text = String::new();
    let mut in_text = false;

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element))
            | Ok(quick_xml::events::Event::Empty(ref element)) => {
                match element.local_name().as_ref() {
                    b"footnote" | b"endnote" => {
                        if let Some(id) = current_id.take() {
                            let text = current_text.trim().to_string();
                            if !text.is_empty() {
                                map.insert(id, text);
                            }
                        }
                        current_text.clear();
                        for attribute in element.attributes().flatten() {
                            if attribute.key.local_name().as_ref() == b"id"
                                && let Ok(value) = attribute.unescape_value()
                            {
                                current_id = value.parse::<usize>().ok();
                            }
                        }
                    }
                    b"t" => in_text = true,
                    _ => {}
                }
            }
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
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
            },
            Ok(quick_xml::events::Event::Text(ref element)) => {
                if in_text && let Ok(text) = element.xml_content() {
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

fn scan_note_refs(xml: &str) -> Vec<(NoteKind, usize)> {
    let mut refs: Vec<(NoteKind, usize)> = Vec::new();
    let mut reader = quick_xml::Reader::from_str(xml);

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element))
            | Ok(quick_xml::events::Event::Empty(ref element)) => {
                let kind = match element.local_name().as_ref() {
                    b"footnoteReference" => Some(NoteKind::Footnote),
                    b"endnoteReference" => Some(NoteKind::Endnote),
                    _ => None,
                };
                if let Some(kind) = kind {
                    for attribute in element.attributes().flatten() {
                        if attribute.key.local_name().as_ref() == b"id"
                            && let Ok(value) = attribute.unescape_value()
                            && let Ok(id) = value.parse::<usize>()
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

pub(super) fn is_note_reference_run(run: &docx_rs::Run, notes: &NoteContext) -> bool {
    if let Some(ref style) = run.run_property.style
        && notes.note_style_ids.contains(&style.val)
    {
        return extract_run_text(run).is_empty();
    }
    false
}

#[cfg(test)]
mod paragraph_container_tests {
    use super::ParagraphContainerContext;
    use crate::ir::{BorderLineStyle, Color, Insets, ParagraphBorderSide};

    #[test]
    fn test_paragraph_container_context_extracts_background_border_and_padding() {
        let xml = r#"
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:r><w:t>Plain</w:t></w:r>
                </w:p>
                <w:p>
                  <w:pPr>
                    <w:pBdr>
                      <w:top w:val="single" w:sz="6" w:space="10" w:color="E1E4E8"/>
                      <w:left w:val="single" w:sz="6" w:space="10" w:color="E1E4E8"/>
                      <w:bottom w:val="single" w:sz="6" w:space="10" w:color="E1E4E8"/>
                      <w:right w:val="single" w:sz="6" w:space="10" w:color="E1E4E8"/>
                    </w:pBdr>
                    <w:shd w:fill="F6F8FA"/>
                  </w:pPr>
                  <w:r><w:t>Code</w:t></w:r>
                </w:p>
              </w:body>
            </w:document>
        "#;

        let context = ParagraphContainerContext::from_xml(Some(xml));
        assert_eq!(context.next_style(), None);

        let container = context
            .next_style()
            .expect("second paragraph should produce a container style");
        assert_eq!(container.background, Some(Color::new(246, 248, 250)));
        assert_eq!(
            container.padding,
            Some(Insets {
                top: 10.0,
                right: 10.0,
                bottom: 10.0,
                left: 10.0,
            })
        );

        let border = container.border.expect("expected paragraph border");
        let expected_side = ParagraphBorderSide {
            width: 0.75,
            color: Color::new(225, 228, 232),
            style: BorderLineStyle::Solid,
        };
        assert_eq!(border.top, Some(expected_side.clone()));
        assert_eq!(border.right, Some(expected_side.clone()));
        assert_eq!(border.bottom, Some(expected_side.clone()));
        assert_eq!(border.left, Some(expected_side));
    }
}
