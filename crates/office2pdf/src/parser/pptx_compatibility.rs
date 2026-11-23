use std::borrow::Cow;

use quick_xml::events::Event;
use quick_xml::name::{QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::ConvertError;

const MC: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Select one representation before any shape, chart or inheritance scan.
/// Keep selected wrappers so their namespace bindings remain in scope, and
/// copy original XML slices so text whitespace and entities remain unchanged.
pub(super) fn select_alternate_content(xml: &str) -> Result<Cow<'_, str>, ConvertError> {
    if !xml.contains("AlternateContent") {
        return Ok(Cow::Borrowed(xml));
    }
    let started = std::time::Instant::now();
    let mut reader = NsReader::from_str(xml);
    let mut scopes = Vec::new();
    let mut depth = 0usize;
    let mut kept_from = 0usize;
    let mut output = String::new();
    let mut omitted = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|error| {
            crate::parser::parse_err(format!("XML error in alternate content: {error}"))
        })?;
        match &event {
            Event::Start(element) | Event::Empty(element) => {
                let empty = matches!(event, Event::Empty(_));
                let (namespace, name) = reader.resolver().resolve_element(element.name());
                let is_mc = matches!(namespace, ResolveResult::Bound(ns) if ns.as_ref() == MC);
                if is_mc && name.as_ref() == b"AlternateContent" && !empty {
                    scopes.push((depth, false));
                } else if is_mc
                    && matches!(name.as_ref(), b"Choice" | b"Fallback")
                    && let Some((scope_depth, selected)) = scopes.last_mut()
                    && depth == *scope_depth + 1
                {
                    let supported = name.as_ref() == b"Fallback"
                        || super::get_attr_str(element, b"Requires").is_some_and(|requires| {
                            let mut prefixes = requires.split_ascii_whitespace().peekable();
                            prefixes.peek().is_some()
                                && prefixes.all(|prefix| {
                                    let qualified = format!("{prefix}:required");
                                    matches!(reader.resolver().resolve_element(QName(qualified.as_bytes())).0,
                                        ResolveResult::Bound(ns) if supported_namespace(ns.as_ref()))
                                })
                        });
                    if *selected || !supported {
                        if !empty {
                            reader.read_to_end(element.name()).map_err(|error| {
                                crate::parser::parse_err(format!(
                                    "XML error in alternate branch: {error}"
                                ))
                            })?;
                        }
                        output.push_str(&xml[kept_from..start]);
                        kept_from = reader.buffer_position() as usize;
                        omitted += 1;
                        continue;
                    }
                    *selected = true;
                }
                if !empty {
                    depth += 1;
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if scopes
                    .last()
                    .is_some_and(|&(scope_depth, _)| scope_depth == depth)
                {
                    scopes.pop();
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    tracing::debug!(
        omitted_branches = omitted,
        elapsed_ms = started.elapsed().as_millis(),
        "selected PPTX alternate representations"
    );
    if omitted == 0 {
        Ok(Cow::Borrowed(xml))
    } else {
        output.push_str(&xml[kept_from..]);
        Ok(Cow::Owned(output))
    }
}

/// Namespace URIs, rather than conventional prefixes, identify the drawing
/// vocabularies consumed by the slide parser. Unknown extensions use fallback.
fn supported_namespace(namespace: &[u8]) -> bool {
    matches!(
        namespace,
        b"http://schemas.openxmlformats.org/drawingml/2006/main"
            | b"http://schemas.openxmlformats.org/drawingml/2006/chart"
            | b"http://schemas.openxmlformats.org/drawingml/2006/diagram"
            | b"http://schemas.openxmlformats.org/presentationml/2006/main"
            | b"http://schemas.openxmlformats.org/officeDocument/2006/math"
            | b"http://schemas.microsoft.com/office/drawing/2010/main"
    )
}
