use std::io::Read;

use quick_xml::Reader;
use quick_xml::events::Event;
use zip::ZipArchive;

use crate::ir::Metadata;

/// Parse Dublin Core metadata from OOXML `docProps/core.xml` inside a ZIP archive.
///
/// Returns `Metadata::default()` if the entry is missing or unparseable (no error).
pub fn extract_metadata_from_zip<R: Read + std::io::Seek>(archive: &mut ZipArchive<R>) -> Metadata {
    let xml = match archive.by_name("docProps/core.xml") {
        Ok(mut file) => {
            let mut content = String::new();
            if file.read_to_string(&mut content).is_err() {
                return Metadata::default();
            }
            content
        }
        Err(_) => return Metadata::default(),
    };
    parse_core_xml(&xml)
}

/// Parse Dublin Core metadata from `docProps/core.xml` content string.
///
/// Extracts: `dc:title`, `dc:creator`, `dc:subject`, `dc:description`,
/// `dcterms:created`, `dcterms:modified`.
pub fn parse_core_xml(xml: &str) -> Metadata {
    let mut metadata = Metadata::default();
    let mut reader = Reader::from_str(xml);

    #[derive(Clone, Copy, PartialEq, Eq)]
    enum Field {
        None,
        Title,
        Creator,
        Subject,
        Description,
        Created,
        Modified,
    }

    let mut current = Field::None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let local = e.local_name();
                current = match local.as_ref() {
                    b"title" => Field::Title,
                    b"creator" => Field::Creator,
                    b"subject" => Field::Subject,
                    b"description" => Field::Description,
                    b"created" => Field::Created,
                    b"modified" => Field::Modified,
                    _ => Field::None,
                };
            }
            Ok(Event::Text(e)) => {
                if current != Field::None
                    && let Ok(text) = e.xml_content()
                {
                    let text = text.to_string();
                    if !text.is_empty() {
                        match current {
                            Field::Title => metadata.title = Some(text),
                            Field::Creator => metadata.author = Some(text),
                            Field::Subject => metadata.subject = Some(text),
                            Field::Description => metadata.description = Some(text),
                            Field::Created => metadata.created = Some(text),
                            Field::Modified => metadata.modified = Some(text),
                            Field::None => {}
                        }
                    }
                }
            }
            Ok(Event::End(_)) => {
                current = Field::None;
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    metadata
}

#[cfg(test)]
#[path = "metadata_tests.rs"]
mod tests;
