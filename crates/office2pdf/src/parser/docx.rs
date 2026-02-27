use std::cell::Cell;
use std::collections::HashMap;
use std::io::Read;

use crate::config::ConvertOptions;
use crate::error::{ConvertError, ConvertWarning};
use crate::ir::{
    Alignment, Block, BorderSide, CellBorder, Chart, Color, Document, FloatingImage, FlowPage,
    HFInline, HeaderFooter, HeaderFooterParagraph, ImageData, ImageFormat, LineSpacing, List,
    ListItem, ListKind, Margins, MathEquation, Metadata, Page, PageSize, Paragraph, ParagraphStyle,
    Run, StyleSheet, Table, TableCell, TableRow, TextStyle, WrapMode,
};
use crate::parser::Parser;

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

/// Numbering info extracted from a paragraph's numPr.
#[derive(Debug, Clone)]
struct NumInfo {
    num_id: usize,
    level: u32,
}

/// Map from numId → ListKind (Ordered or Unordered).
/// Built by resolving numId → abstractNumId → first level's format.
type NumKindMap = HashMap<usize, ListKind>;

/// Build a map from numbering instance ID → list kind by inspecting the
/// abstract numbering definitions.
fn build_num_kind_map(numberings: &docx_rs::Numberings) -> NumKindMap {
    // Map abstractNumId → is bullet?
    let mut abstract_kinds: HashMap<usize, ListKind> = HashMap::new();
    for abs in &numberings.abstract_nums {
        // Check the first level's format to determine if bullet or ordered
        let kind = if abs.levels.iter().any(|lvl| {
            let json = serde_json::to_value(&lvl.format).ok();
            json.and_then(|j| j.as_str().map(|s| s.to_owned()))
                .is_some_and(|val| val == "bullet")
        }) {
            ListKind::Unordered
        } else {
            ListKind::Ordered
        };
        abstract_kinds.insert(abs.id, kind);
    }

    // Map numId → abstractNumId → ListKind
    let mut map = NumKindMap::new();
    for num in &numberings.numberings {
        if let Some(&kind) = abstract_kinds.get(&num.abstract_num_id) {
            map.insert(num.id, kind);
        }
    }
    map
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
    /// Heading level from outline_lvl (0 = Heading 1, 1 = Heading 2, ..., 5 = Heading 6).
    heading_level: Option<usize>,
}

/// Map from style_id → resolved formatting.
type StyleMap = HashMap<String, ResolvedStyle>;

/// Default font sizes for heading levels (Heading 1-6).
/// Index 0 = Heading 1 (outline_lvl 0), index 5 = Heading 6 (outline_lvl 5).
const HEADING_DEFAULT_SIZES: [f64; 6] = [24.0, 20.0, 16.0, 14.0, 12.0, 11.0];

/// Build a map from style ID → resolved formatting by extracting formatting
/// from each style's run_property and paragraph_property.
fn build_style_map(styles: &docx_rs::Styles) -> StyleMap {
    let mut map = StyleMap::new();
    for style in &styles.styles {
        // Only process paragraph styles (not character or table styles)
        if style.style_type != docx_rs::StyleType::Paragraph {
            continue;
        }

        let text = extract_run_style(&style.run_property);
        let paragraph = extract_paragraph_style(&style.paragraph_property);
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

    merged
}

/// Merge style paragraph formatting with explicit paragraph formatting.
/// Explicit formatting takes priority.
fn merge_paragraph_style(
    explicit: &ParagraphStyle,
    style: Option<&ResolvedStyle>,
) -> ParagraphStyle {
    let style_para = match style {
        Some(s) => &s.paragraph,
        None => return explicit.clone(),
    };

    ParagraphStyle {
        alignment: explicit.alignment.or(style_para.alignment),
        indent_left: explicit.indent_left.or(style_para.indent_left),
        indent_right: explicit.indent_right.or(style_para.indent_right),
        indent_first_line: explicit.indent_first_line.or(style_para.indent_first_line),
        line_spacing: explicit.line_spacing.or(style_para.line_spacing),
        space_before: explicit.space_before.or(style_para.space_before),
        space_after: explicit.space_after.or(style_para.space_after),
    }
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

/// Group consecutive list paragraphs (with the same numId) into List blocks.
/// Non-list elements pass through unchanged.
fn group_into_lists(elements: Vec<TaggedElement>, num_kinds: &NumKindMap) -> Vec<Block> {
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
                        });
                        continue;
                    }
                    // Different list — flush current
                    let kind = num_kinds
                        .get(&cur_num_id)
                        .copied()
                        .unwrap_or(ListKind::Unordered);
                    result.push(Block::List(List {
                        kind,
                        items: std::mem::take(items),
                    }));
                }
                // Start new list
                current_list = Some((
                    info.num_id,
                    vec![ListItem {
                        content: vec![paragraph],
                        level: info.level,
                    }],
                ));
            }
            TaggedElement::Plain(blocks) => {
                // Flush any pending list
                if let Some((num_id, items)) = current_list.take() {
                    let kind = num_kinds
                        .get(&num_id)
                        .copied()
                        .unwrap_or(ListKind::Unordered);
                    result.push(Block::List(List { kind, items }));
                }
                result.extend(blocks);
            }
        }
    }

    // Flush trailing list
    if let Some((num_id, items)) = current_list {
        let kind = num_kinds
            .get(&num_id)
            .copied()
            .unwrap_or(ListKind::Unordered);
        result.push(Block::List(List { kind, items }));
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
}

/// Wrap type info for an anchor image, scanned from raw document XML.
struct AnchorWrapInfo {
    wrap_mode: WrapMode,
    behind_doc: bool,
}

/// Context for resolving wrap modes of anchor images during parsing.
/// The `cursor` is advanced each time an anchor image is encountered.
struct WrapContext {
    /// Ordered list of wrap info for anchor images as they appear in document.xml.
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

/// Build a `WrapContext` by scanning document.xml for anchor wrap types.
/// docx-rs does not parse wrap type elements (wrapSquare, wrapNone, etc.),
/// so we scan the raw XML like we do for footnotes.
fn build_wrap_context(data: &[u8]) -> WrapContext {
    let mut wraps = Vec::new();

    let Ok(mut archive) = zip::ZipArchive::new(std::io::Cursor::new(data)) else {
        return WrapContext {
            wraps,
            cursor: Cell::new(0),
        };
    };

    if let Some(xml) = read_zip_text(&mut archive, "word/document.xml") {
        wraps = scan_anchor_wrap_types(&xml);
    }

    WrapContext {
        wraps,
        cursor: Cell::new(0),
    }
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

/// Build a `MathContext` by scanning document.xml for OMML elements.
fn build_math_context(data: &[u8]) -> MathContext {
    let mut equations: HashMap<usize, Vec<MathEquation>> = HashMap::new();

    let Ok(mut archive) = zip::ZipArchive::new(std::io::Cursor::new(data)) else {
        return MathContext { equations };
    };

    if let Some(xml) = read_zip_text(&mut archive, "word/document.xml") {
        let raw = super::omml::scan_math_equations(&xml);
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

/// Build a `ChartContext` by scanning document.xml for chart references
/// and parsing the corresponding chart XML files from the ZIP.
fn build_chart_context(data: &[u8]) -> ChartContext {
    let mut charts: HashMap<usize, Vec<Chart>> = HashMap::new();

    let Ok(mut archive) = zip::ZipArchive::new(std::io::Cursor::new(data)) else {
        return ChartContext { charts };
    };

    // Read document.xml and relationships file
    let doc_xml = read_zip_text(&mut archive, "word/document.xml");
    let rels_xml = read_zip_text(&mut archive, "word/_rels/document.xml.rels");

    let (Some(doc_xml), Some(rels_xml)) = (doc_xml, rels_xml) else {
        return ChartContext { charts };
    };

    // Find chart references in document.xml
    let refs = super::chart::scan_chart_references(&doc_xml);
    // Build relationship ID → chart file path mapping
    let rels = super::chart::scan_chart_rels(&rels_xml);

    // For each chart reference, parse the chart XML
    for (body_idx, rid) in refs {
        if let Some(chart_path) = rels.get(&rid)
            && let Some(chart_xml) = read_zip_text(&mut archive, chart_path)
            && let Some(chart) = super::chart::parse_chart_xml(&chart_xml)
        {
            charts.entry(body_idx).or_default().push(chart);
        }
    }

    ChartContext { charts }
}

/// Build a `NoteContext` by parsing footnotes/endnotes from the raw DOCX ZIP.
/// This is needed because docx-rs reader does not parse `w:footnoteReference`
/// or `w:endnoteReference` elements.
fn build_note_context(data: &[u8]) -> NoteContext {
    let mut footnote_content = HashMap::new();
    let mut endnote_content = HashMap::new();
    let mut note_refs = Vec::new();

    let Ok(mut archive) = zip::ZipArchive::new(std::io::Cursor::new(data)) else {
        return NoteContext {
            footnote_content,
            endnote_content,
            note_refs,
            cursor: Cell::new(0),
        };
    };

    // Parse word/footnotes.xml
    if let Some(xml) = read_zip_text(&mut archive, "word/footnotes.xml") {
        footnote_content = parse_notes_xml(&xml);
    }

    // Parse word/endnotes.xml
    if let Some(xml) = read_zip_text(&mut archive, "word/endnotes.xml") {
        endnote_content = parse_notes_xml(&xml);
    }

    // Scan word/document.xml for note references in document order
    if let Some(xml) = read_zip_text(&mut archive, "word/document.xml") {
        note_refs = scan_note_refs(&xml);
    }

    NoteContext {
        footnote_content,
        endnote_content,
        note_refs,
        cursor: Cell::new(0),
    }
}

/// Read a ZIP entry as a UTF-8 string.
fn read_zip_text(
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
    name: &str,
) -> Option<String> {
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
/// These runs have `rStyle` set to "FootnoteReference" or "EndnoteReference"
/// and contain no text.
fn is_note_reference_run(run: &docx_rs::Run) -> bool {
    if let Some(ref style) = run.run_property.style {
        let val = &style.val;
        if val == "FootnoteReference" || val == "EndnoteReference" {
            // Verify the run has no text content (only footnoteReference element)
            return extract_run_text(run).is_empty();
        }
    }
    false
}

impl Parser for DocxParser {
    fn parse(
        &self,
        data: &[u8],
        _options: &ConvertOptions,
    ) -> Result<(Document, Vec<ConvertWarning>), ConvertError> {
        // Build note context from raw ZIP before docx-rs parsing
        let notes = build_note_context(data);
        // Build wrap context for anchor image wrap types from raw ZIP
        let wraps = build_wrap_context(data);
        // Build math context for OMML equations from raw ZIP
        let mut math = build_math_context(data);
        // Build chart context for embedded charts from raw ZIP
        let mut chart_ctx = build_chart_context(data);

        let docx = docx_rs::read_docx(data)
            .map_err(|e| ConvertError::Parse(format!("Failed to parse DOCX: {e}")))?;

        let (size, margins) = extract_page_setup(&docx.document.section_property);
        let images = build_image_map(&docx);
        let hyperlinks = build_hyperlink_map(&docx);
        let num_kinds = build_num_kind_map(&docx.numberings);
        let style_map = build_style_map(&docx.styles);
        let mut warnings = Vec::new();

        let mut elements: Vec<TaggedElement> = Vec::new();
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
                    ))])]
                }
                docx_rs::DocumentChild::StructuredDataTag(sdt) => {
                    convert_sdt_children(sdt, &images, &hyperlinks, &style_map, &notes, &wraps)
                }
                _ => vec![TaggedElement::Plain(vec![])],
            }));

            match result {
                Ok(elems) => elements.extend(elems),
                Err(_) => {
                    warnings.push(ConvertWarning {
                        element: format!("Document element at index {idx}"),
                        reason: "element processing panicked; skipped".to_string(),
                    });
                }
            }
        }

        let content = group_into_lists(elements, &num_kinds);

        let header = extract_docx_header(&docx.document.section_property);
        let footer = extract_docx_footer(&docx.document.section_property);

        Ok((
            Document {
                metadata: Metadata::default(),
                pages: vec![Page::Flow(FlowPage {
                    size,
                    margins,
                    content,
                    header,
                    footer,
                })],
                styles: StyleSheet::default(),
            },
            warnings,
        ))
    }
}

/// Extract the default header from DOCX section properties, if present.
fn extract_docx_header(section_prop: &docx_rs::SectionProperty) -> Option<HeaderFooter> {
    let (_rid, header) = section_prop.header.as_ref()?;
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

/// Extract the default footer from DOCX section properties, if present.
fn extract_docx_footer(section_prop: &docx_rs::SectionProperty) -> Option<HeaderFooter> {
    let (_rid, footer) = section_prop.footer.as_ref()?;
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

/// Convert a docx-rs Paragraph into a HeaderFooterParagraph.
/// Detects PAGE field codes within runs and emits HFInline::PageNumber.
fn convert_hf_paragraph(para: &docx_rs::Paragraph) -> HeaderFooterParagraph {
    let style = extract_paragraph_style(&para.property);
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
/// Recognizes text, tabs, and PAGE field codes.
fn extract_hf_run_elements(
    children: &[docx_rs::RunChild],
    style: &TextStyle,
    elements: &mut Vec<HFInline>,
) {
    let mut in_field = false;
    let mut field_is_page = false;
    let mut past_separate = false;

    for child in children {
        match child {
            docx_rs::RunChild::FieldChar(fc) => match fc.field_char_type {
                docx_rs::FieldCharType::Begin => {
                    in_field = true;
                    field_is_page = false;
                    past_separate = false;
                }
                docx_rs::FieldCharType::Separate => {
                    past_separate = true;
                }
                docx_rs::FieldCharType::End => {
                    if field_is_page {
                        elements.push(HFInline::PageNumber);
                    }
                    in_field = false;
                    field_is_page = false;
                    past_separate = false;
                }
                _ => {}
            },
            docx_rs::RunChild::InstrText(instr) => {
                if in_field && matches!(instr.as_ref(), docx_rs::InstrText::PAGE(_)) {
                    field_is_page = true;
                }
            }
            docx_rs::RunChild::InstrTextString(s) => {
                // After round-tripping through build/read_docx, InstrText::PAGE
                // becomes InstrTextString("PAGE").
                if in_field && s.trim().eq_ignore_ascii_case("page") {
                    field_is_page = true;
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
fn convert_sdt_children(
    sdt: &docx_rs::StructuredDataTag,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
) -> Vec<TaggedElement> {
    let mut result = Vec::new();
    for child in &sdt.children {
        match child {
            docx_rs::StructuredDataTagChild::Paragraph(para) => {
                result.push(convert_paragraph_element(
                    para, images, hyperlinks, style_map, notes, wraps,
                ));
            }
            docx_rs::StructuredDataTagChild::Table(table) => {
                result.push(TaggedElement::Plain(vec![Block::Table(convert_table(
                    table, images, hyperlinks, style_map, notes, wraps,
                ))]));
            }
            docx_rs::StructuredDataTagChild::StructuredDataTag(nested) => {
                result.extend(convert_sdt_children(
                    nested, images, hyperlinks, style_map, notes, wraps,
                ));
            }
            _ => {}
        }
    }
    result
}

/// Convert a docx-rs Paragraph into a TaggedElement.
/// If the paragraph has numbering, returns a `ListParagraph`; otherwise `Plain`.
fn convert_paragraph_element(
    para: &docx_rs::Paragraph,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
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
fn convert_paragraph_blocks(
    para: &docx_rs::Paragraph,
    out: &mut Vec<Block>,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
) {
    // Emit page break before the paragraph if requested
    if para.property.page_break_before == Some(true) {
        out.push(Block::PageBreak);
    }

    // Look up the paragraph's referenced style
    let resolved_style = get_paragraph_style_id(&para.property).and_then(|id| style_map.get(id));

    // Collect text runs and detect inline images
    let mut runs: Vec<Run> = Vec::new();
    let mut inline_images: Vec<Block> = Vec::new();

    for child in &para.children {
        match child {
            docx_rs::ParagraphChild::Run(run) => {
                // Check for footnote/endnote reference runs
                if is_note_reference_run(run) {
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

                // Check for images in this run (inline or floating)
                for run_child in &run.children {
                    if let docx_rs::RunChild::Drawing(drawing) = run_child
                        && let Some(img_block) = extract_drawing_image(drawing, images, wraps)
                    {
                        inline_images.push(img_block);
                    }
                }

                // Extract text from the run
                let text = extract_run_text(run);
                if !text.is_empty() {
                    let explicit_style = extract_run_style(&run.run_property);
                    runs.push(Run {
                        text,
                        style: merge_text_style(&explicit_style, resolved_style),
                        href: None,
                        footnote: None,
                    });
                }
            }
            docx_rs::ParagraphChild::Hyperlink(hyperlink) => {
                // Resolve the hyperlink URL from document relationships
                let href = resolve_hyperlink_url(hyperlink, hyperlinks);

                // Extract runs from inside the hyperlink element
                for hchild in &hyperlink.children {
                    if let docx_rs::ParagraphChild::Run(run) = hchild {
                        let text = extract_run_text(run);
                        if !text.is_empty() {
                            let explicit_style = extract_run_style(&run.run_property);
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

    let explicit_para_style = extract_paragraph_style(&para.property);
    out.push(Block::Paragraph(Paragraph {
        style: merge_paragraph_style(&explicit_para_style, resolved_style),
        runs,
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

    ParagraphStyle {
        alignment,
        indent_left,
        indent_right,
        indent_first_line,
        line_spacing,
        space_before,
        space_after,
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
) -> Table {
    let column_widths: Vec<f64> = table.grid.iter().map(|&w| w as f64 / 20.0).collect();

    // First pass: extract raw rows with vmerge info for rowspan calculation
    let raw_rows = extract_raw_rows(table, images, hyperlinks, style_map, notes, wraps);

    // Second pass: resolve vertical merges into rowspan values and build IR rows
    let rows = resolve_vmerge_and_build_rows(&raw_rows);

    Table {
        rows,
        column_widths,
    }
}

/// Intermediate cell representation for vmerge resolution.
struct RawCell {
    content: Vec<Block>,
    col_span: u32,
    col_index: usize,
    vmerge: Option<String>, // "restart", "continue", or None
    border: Option<CellBorder>,
    background: Option<Color>,
}

/// Extract raw rows from a docx-rs Table, tracking column indices and vmerge state.
fn extract_raw_rows(
    table: &docx_rs::Table,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
) -> Vec<Vec<RawCell>> {
    let mut raw_rows = Vec::new();

    for table_child in &table.rows {
        let docx_rs::TableChild::TableRow(row) = table_child;
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

            let content = extract_cell_content(cell, images, hyperlinks, style_map, notes, wraps);
            let border = prop_json
                .as_ref()
                .and_then(|j| j.get("borders"))
                .and_then(extract_cell_borders);
            let background = prop_json
                .as_ref()
                .and_then(|j| j.get("shading"))
                .and_then(extract_cell_shading);

            cells.push(RawCell {
                content,
                col_span: grid_span,
                col_index,
                vmerge,
                border,
                background,
            });

            col_index += grid_span as usize;
        }

        raw_rows.push(cells);
    }

    raw_rows
}

/// Resolve vertical merges: compute rowspan for "restart" cells and skip "continue" cells.
fn resolve_vmerge_and_build_rows(raw_rows: &[Vec<RawCell>]) -> Vec<TableRow> {
    let mut rows = Vec::new();

    for (row_idx, raw_row) in raw_rows.iter().enumerate() {
        let mut cells = Vec::new();

        for raw_cell in raw_row {
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
                    });
                }
            }
        }

        rows.push(TableRow {
            cells,
            height: None,
        });
    }

    rows
}

/// Count the vertical merge span starting from a "restart" cell.
/// Looks at rows below `start_row` for "continue" cells at the same column index.
fn count_vmerge_span(raw_rows: &[Vec<RawCell>], start_row: usize, col_index: usize) -> u32 {
    let mut span = 1u32;
    for row in raw_rows.iter().skip(start_row + 1) {
        let has_continue = row
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

/// Extract cell content (paragraphs) from a docx-rs TableCell.
fn extract_cell_content(
    cell: &docx_rs::TableCell,
    images: &ImageMap,
    hyperlinks: &HyperlinkMap,
    style_map: &StyleMap,
    notes: &NoteContext,
    wraps: &WrapContext,
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
                );
            }
            docx_rs::TableCellContent::Table(nested_table) => {
                blocks.push(Block::Table(convert_table(
                    nested_table,
                    images,
                    hyperlinks,
                    style_map,
                    notes,
                    wraps,
                )));
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
        Some(BorderSide {
            width: size / 8.0, // eighths of a point → points
            color,
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
/// docx-rs types with private fields serialize directly as their inner value
/// (e.g. Bold → `true`, Sz → `24`, Color → `"FF0000"`), not as `{"val": ...}`.
/// Strike has a public `val` field and can be accessed directly.
fn extract_run_style(rp: &docx_rs::RunProperty) -> TextStyle {
    TextStyle {
        bold: extract_bool_prop(&rp.bold),
        italic: extract_bool_prop(&rp.italic),
        underline: rp.underline.as_ref().and_then(|u| {
            let json = serde_json::to_value(u).ok()?;
            let val = json.as_str()?;
            if val == "none" { None } else { Some(true) }
        }),
        strikethrough: rp.strike.as_ref().map(|s| s.val),
        font_size: rp.sz.as_ref().and_then(|sz| {
            let json = serde_json::to_value(sz).ok()?;
            let half_points = json.as_f64()?;
            Some(half_points / 2.0)
        }),
        color: rp.color.as_ref().and_then(|c| {
            let json = serde_json::to_value(c).ok()?;
            let hex = json.as_str()?;
            parse_hex_color(hex)
        }),
        font_family: rp.fonts.as_ref().and_then(|f| {
            let json = serde_json::to_value(f).ok()?;
            // Prefer ascii font name, fall back to hi_ansi, east_asia, cs
            json.get("ascii")
                .or_else(|| json.get("hi_ansi"))
                .or_else(|| json.get("east_asia"))
                .or_else(|| json.get("cs"))
                .and_then(|v| v.as_str())
                .map(String::from)
        }),
    }
}

/// Extract a boolean property (Bold, Italic) via serde. Returns None if absent.
/// docx-rs serializes Bold/Italic directly as a boolean (e.g. `true`).
fn extract_bool_prop<T: serde::Serialize>(prop: &Option<T>) -> Option<bool> {
    prop.as_ref().and_then(|p| {
        let json = serde_json::to_value(p).ok()?;
        json.as_bool()
    })
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
        assert_eq!(math_blocks[0].content, "E=mc^2");
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
}
