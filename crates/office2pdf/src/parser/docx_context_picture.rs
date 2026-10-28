use std::cell::Cell;

use quick_xml::NsReader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;

use crate::ir::{FrameAlign, FrameAnchor, ImageCrop, Insets, WrapMode};
use crate::parser::xml_util::OOXML_XML_VERSION;

const WORD: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PICTURE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/picture";
const DRAWING: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const WORD_DRAWING: &[u8] =
    b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

/// What one `w:drawing` states that the released docx-rs does not expose.
#[derive(Clone, Copy, Default)]
pub(in super::super) struct DrawingMetadata {
    pub(in super::super) crop: Option<ImageCrop>,
    /// A floating drawing whose `wp:docPr` states `hidden`. Word neither
    /// draws it nor wraps text around it (issue #2299; LibreOffice tdf#119201,
    /// tdf#162527 and tdf#169802 compare Word against Writer).
    /// TODO(no Word evidence for `wp:inline`): whether a hidden inline picture
    /// still reserves its line height is unknown, so only `wp:anchor` is read.
    pub(in super::super) hidden: bool,
    /// Where a `wp:anchor` sits and how text makes room for it; `None` for a
    /// `wp:inline` drawing. docx-rs reads an unknown `relativeFrom` as
    /// `margin` and keeps no wrap, so a story picture is placed from these
    /// (issue #2327).
    pub(in super::super) anchor: Option<AnchorPlacement>,
}

/// A `wp:anchor`'s position and text wrap, as its own children state them.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub(in super::super) struct AnchorPlacement {
    pub(in super::super) horizontal: AnchorAxisPlacement,
    pub(in super::super) vertical: AnchorAxisPlacement,
    /// The `wp:wrap*` element; `None` where the anchor states none.
    wrap: Option<WrapMode>,
    behind_doc: bool,
    /// `distT`, `distR`, `distB` and `distL`, in points.
    pub(in super::super) distance: Insets,
}

impl AnchorPlacement {
    /// How text makes room for the drawing, read as the body's wrap scan
    /// reads it: `behindDoc` puts it behind whatever wrap it states.
    pub(in super::super) fn wrap_mode(&self) -> WrapMode {
        if self.behind_doc {
            WrapMode::Behind
        } else {
            self.wrap.unwrap_or(WrapMode::None)
        }
    }
}

/// One `wp:positionH` or `wp:positionV`.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub(in super::super) struct AnchorAxisPlacement {
    /// `relativeFrom` when it names the page, its margins or (horizontally)
    /// the column; `None` for the frames that move with the text (paragraph,
    /// line, character) and the single-margin or odd/even ones.
    pub(in super::super) reference: Option<FrameAnchor>,
    /// `wp:posOffset`, in points.
    pub(in super::super) offset_pt: Option<f64>,
    /// `wp:align`; `inside` and `outside` depend on the page and stay `None`.
    pub(in super::super) align: Option<FrameAlign>,
}

#[derive(Default)]
struct ScannedDrawing {
    metadata: DrawingMetadata,
    subtree_end: usize,
}

/// The released docx-rs does not expose source rectangles, `wp:docPr`
/// visibility or a `wp:anchor`'s wrap. Keep metadata per drawing occurrence,
/// not per media relationship, which pictures can share.
#[derive(Default)]
pub(in super::super) struct PictureCropContext {
    drawings: Vec<ScannedDrawing>,
    cursor: Cell<usize>,
}

impl PictureCropContext {
    pub(in super::super) fn from_xml(xml: Option<&str>) -> Self {
        Self {
            drawings: xml.map(scan_drawings).unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    /// Header text-box children use a separate conversion path;
    /// skip their nested drawing entries when the ordinary path does not visit them.
    pub(in super::super) fn next_drawing(&self, skip_nested: bool) -> DrawingMetadata {
        let index: usize = self.cursor.get();
        let Some(drawing) = self.drawings.get(index) else {
            return DrawingMetadata::default();
        };
        self.cursor.set(if skip_nested {
            drawing.subtree_end
        } else {
            index + 1
        });
        drawing.metadata
    }
}

fn namespace(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Vec<u8> {
    match reader.resolver().resolve_element(element.name()).0 {
        ResolveResult::Bound(value) => value.as_ref().to_vec(),
        _ => Vec::new(),
    }
}

fn source_rectangle(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Option<ImageCrop> {
    let mut crop: ImageCrop = ImageCrop::default();
    for attribute in element.attributes().flatten() {
        // Source-rectangle attributes are unqualified; alien namespaced
        // attributes must not override the four DrawingML values.
        if attribute.key.prefix().is_some() {
            continue;
        }
        let Ok(value) = attribute.decoded_and_normalized_value(OOXML_XML_VERSION, reader.decoder())
        else {
            continue;
        };
        let value: &str = value.trim();
        let offset: Option<f64> = if let Some(percent) = value.strip_suffix('%') {
            percent
                .trim()
                .parse::<f64>()
                .ok()
                .map(|number| number / 100.0)
        } else {
            value
                .parse::<i32>()
                .ok()
                .map(|number| f64::from(number) / 100_000.0)
        };
        let Some(offset) = offset.filter(|number| number.is_finite()) else {
            continue;
        };
        match attribute.key.as_ref() {
            b"l" => crop.left = offset,
            b"t" => crop.top = offset,
            b"r" => crop.right = offset,
            b"b" => crop.bottom = offset,
            _ => {}
        }
    }
    (!crop.is_empty()).then_some(crop)
}

/// Whether a `wp:docPr` states its unqualified `hidden` as an `xsd:boolean`
/// true.
pub(in super::super) fn doc_pr_states_hidden(element: &BytesStart<'_>) -> bool {
    element.attributes().flatten().any(|attribute| {
        attribute.key.as_ref() == b"hidden"
            && matches!(attribute.value.trim_ascii(), b"1" | b"true")
    })
}

/// An unqualified attribute's value, decoded.
fn attribute_value(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Option<String> {
    element
        .attributes()
        .flatten()
        .find(|attribute| attribute.key.as_ref() == name)
        .and_then(|attribute| {
            attribute
                .decoded_and_normalized_value(OOXML_XML_VERSION, reader.decoder())
                .ok()
                .map(|value| value.trim().to_string())
        })
}

/// What a `wp:anchor` start tag states: `behindDoc` and the wrap distances.
fn anchor_start(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> AnchorPlacement {
    let distance_pt = |name: &[u8]| -> f64 {
        attribute_value(reader, element, name)
            .and_then(|value| value.parse::<f64>().ok())
            .filter(|emu| emu.is_finite())
            .map_or(0.0, |emu| emu / 12_700.0)
    };
    AnchorPlacement {
        behind_doc: attribute_value(reader, element, b"behindDoc")
            .is_some_and(|value| matches!(value.as_str(), "1" | "true")),
        distance: Insets {
            top: distance_pt(b"distT"),
            right: distance_pt(b"distR"),
            bottom: distance_pt(b"distB"),
            left: distance_pt(b"distL"),
        },
        ..AnchorPlacement::default()
    }
}

fn anchor_reference(value: &str, horizontal: bool) -> Option<FrameAnchor> {
    match value {
        "page" => Some(FrameAnchor::Page),
        "margin" => Some(FrameAnchor::Margin),
        "column" if horizontal => Some(FrameAnchor::Text),
        _ => None,
    }
}

fn anchor_align(value: &str) -> Option<FrameAlign> {
    match value {
        "left" | "top" => Some(FrameAlign::Start),
        "center" => Some(FrameAlign::Center),
        "right" | "bottom" => Some(FrameAlign::End),
        _ => None,
    }
}

fn anchor_wrap_mode(local: &[u8]) -> Option<WrapMode> {
    match local {
        b"wrapNone" => Some(WrapMode::None),
        b"wrapSquare" => Some(WrapMode::Square),
        b"wrapTight" | b"wrapThrough" => Some(WrapMode::Tight),
        b"wrapTopAndBottom" => Some(WrapMode::TopAndBottom),
        _ => None,
    }
}

/// Which `wp:anchor` field the next text node fills.
#[derive(Clone, Copy)]
struct PendingAnchorText {
    drawing: usize,
    horizontal: bool,
    is_align: bool,
}

fn scan_drawings(xml: &str) -> Vec<ScannedDrawing> {
    let mut reader: NsReader<&[u8]> = NsReader::from_str(xml);
    let mut path: Vec<(Vec<u8>, Vec<u8>)> = Vec::new();
    let mut drawings: Vec<ScannedDrawing> = Vec::new();
    let mut open_drawings: Vec<(usize, usize)> = Vec::new();
    let mut skipped_depth: Option<usize> = None;
    let mut pending_text: Option<PendingAnchorText> = None;

    loop {
        let (element, empty) = match reader.read_event() {
            Ok(Event::Start(element)) => (element, false),
            Ok(Event::Empty(element)) => (element, true),
            Ok(Event::Text(text)) => {
                if let Some(pending) = pending_text.take()
                    && let Ok(value) = text.xml_content(OOXML_XML_VERSION)
                    && let Some(anchor) = drawings[pending.drawing].metadata.anchor.as_mut()
                {
                    let axis: &mut AnchorAxisPlacement = if pending.horizontal {
                        &mut anchor.horizontal
                    } else {
                        &mut anchor.vertical
                    };
                    let value: &str = value.trim();
                    if pending.is_align {
                        axis.align = anchor_align(value);
                    } else {
                        axis.offset_pt = value
                            .parse::<f64>()
                            .ok()
                            .filter(|emu| emu.is_finite())
                            .map(|emu| emu / 12_700.0);
                    }
                }
                continue;
            }
            Ok(Event::End(_)) => {
                pending_text = None;
                if open_drawings
                    .last()
                    .is_some_and(|(_, depth)| *depth == path.len())
                {
                    let (index, _) = open_drawings.pop().unwrap();
                    drawings[index].subtree_end = drawings.len();
                }
                if skipped_depth == Some(path.len()) {
                    skipped_depth = None;
                }
                path.pop();
                continue;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => continue,
        };
        let uri: Vec<u8> = namespace(&reader, &element);
        let local: Vec<u8> = element.local_name().as_ref().to_vec();
        let skip: bool = super::paragraph_cursor::is_skipped_subtree(&local)
            || (uri == WORD && matches!(local.as_slice(), b"del" | b"moveFrom" | b"hyperlink"));
        if skipped_depth.is_none() && skip && !empty {
            skipped_depth = Some(path.len() + 1);
        }
        if skipped_depth.is_none() && !skip {
            if uri == WORD && local == b"drawing" {
                let index: usize = drawings.len();
                drawings.push(ScannedDrawing {
                    metadata: DrawingMetadata::default(),
                    subtree_end: index + 1,
                });
                if !empty {
                    open_drawings.push((index, path.len() + 1));
                }
            } else if uri == DRAWING
                && local == b"srcRect"
                && path.len() >= 2
                && path[path.len() - 1].0 == PICTURE
                && path[path.len() - 1].1 == b"blipFill"
                && path[path.len() - 2].0 == PICTURE
                && path[path.len() - 2].1 == b"pic"
                && let Some((index, _)) = open_drawings.last()
            {
                drawings[*index].metadata.crop = source_rectangle(&reader, &element);
            } else if uri == WORD_DRAWING
                && local == b"docPr"
                && path.len() >= 2
                && path[path.len() - 1].0 == WORD_DRAWING
                && path[path.len() - 1].1 == b"anchor"
                && path[path.len() - 2].0 == WORD
                && path[path.len() - 2].1 == b"drawing"
                && let Some((index, _)) = open_drawings.last()
            {
                drawings[*index].metadata.hidden = doc_pr_states_hidden(&element);
            } else if uri == WORD_DRAWING
                && let Some((index, _)) = open_drawings.last()
            {
                scan_anchor_child(
                    &reader,
                    &element,
                    &path,
                    *index,
                    &mut drawings[*index].metadata,
                    &mut pending_text,
                    empty,
                );
            }
        }
        if !empty {
            path.push((uri, local));
        }
    }
    drawings
}

/// Read one `wp:anchor` element, or one of its position and wrap children,
/// onto the open drawing's metadata.
fn scan_anchor_child(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    path: &[(Vec<u8>, Vec<u8>)],
    drawing: usize,
    metadata: &mut DrawingMetadata,
    pending_text: &mut Option<PendingAnchorText>,
    empty: bool,
) {
    let is = |offset: usize, uri: &[u8], local: &[u8]| -> bool {
        path.len() >= offset
            && path[path.len() - offset].0 == uri
            && path[path.len() - offset].1 == local
    };
    let local = element.local_name();
    let local: &[u8] = local.as_ref();
    if local == b"anchor" && is(1, WORD, b"drawing") {
        metadata.anchor = Some(anchor_start(reader, element));
        return;
    }
    let Some(anchor) = metadata.anchor.as_mut() else {
        return;
    };
    if is(1, WORD_DRAWING, b"anchor") && is(2, WORD, b"drawing") {
        if let Some(wrap_mode) = anchor_wrap_mode(local) {
            anchor.wrap = Some(wrap_mode);
        } else if matches!(local, b"positionH" | b"positionV") {
            let horizontal: bool = local == b"positionH";
            let reference: Option<FrameAnchor> = attribute_value(reader, element, b"relativeFrom")
                .and_then(|value| anchor_reference(&value, horizontal));
            if horizontal {
                anchor.horizontal.reference = reference;
            } else {
                anchor.vertical.reference = reference;
            }
        }
    } else if is(2, WORD_DRAWING, b"anchor")
        && is(3, WORD, b"drawing")
        && !empty
        && matches!(local, b"posOffset" | b"align")
    {
        let horizontal: bool = is(1, WORD_DRAWING, b"positionH");
        if horizontal || is(1, WORD_DRAWING, b"positionV") {
            *pending_text = Some(PendingAnchorText {
                drawing,
                horizontal,
                is_align: local == b"align",
            });
        }
    }
}
