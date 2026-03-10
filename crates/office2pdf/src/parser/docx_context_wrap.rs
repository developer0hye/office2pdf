use std::cell::Cell;

use crate::ir::WrapMode;

struct AnchorWrapInfo {
    wrap_mode: WrapMode,
    behind_doc: bool,
}

pub(in super::super) struct WrapContext {
    wraps: Vec<AnchorWrapInfo>,
    cursor: Cell<usize>,
}

impl WrapContext {
    pub(in super::super) fn empty() -> Self {
        Self {
            wraps: Vec::new(),
            cursor: Cell::new(0),
        }
    }

    pub(in super::super) fn consume_next(&self) -> WrapMode {
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

pub(in super::super) fn build_wrap_context_from_xml(doc_xml: Option<&str>) -> WrapContext {
    let wraps = doc_xml.map(scan_anchor_wrap_types).unwrap_or_default();
    WrapContext {
        wraps,
        cursor: Cell::new(0),
    }
}

pub(super) fn extract_vml_wrap_mode_from_element(
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
