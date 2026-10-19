//! Lossless Open Packaging Conventions (OPC) container model.
//!
//! This crate is the package-level foundation for round-trip OOXML editing
//! (issue #746). The existing rendering pipeline in `office2pdf` flattens
//! format-specific information into a draw-only intermediate representation,
//! so it cannot be serialized back into DOCX/PPTX/XLSX without data loss.
//! [`OpcPackage`] instead preserves the original container: every ZIP part —
//! including parts and XML this workspace does not understand — survives a
//! load/save cycle, and untouched part payloads are copied byte-for-byte
//! (raw compressed entry copy, so even the compression choice is kept).
//!
//! Only parts explicitly replaced through [`OpcPackage::set_part_bytes`] are
//! re-encoded on save; [`OpcPackage::dirty_part_names`] reports exactly which
//! parts an edit touched. Relationships and content types are parsed on
//! demand for integrity checks and navigation, always from the authoritative
//! part bytes — never from a lossy projection.

use std::borrow::Cow;
use std::fmt;
use std::io::{Cursor, Read, Write};
use std::path::Path;

use quick_xml::events::{BytesStart, Event};

/// Errors raised while reading, editing, or saving an OPC package.
#[derive(Debug, thiserror::Error)]
pub enum PackageError {
    /// A ZIP container error, including invalid entry metadata.
    #[error("failed to read package container: {0}")]
    Container(#[from] zip::result::ZipError),
    /// The underlying reader failed.
    #[error("I/O failure while accessing package: {0}")]
    Io(#[from] std::io::Error),
    /// A named part does not exist in the package.
    #[error("package has no part named {0:?}")]
    MissingPart(String),
    /// A package part holds XML this crate could not parse.
    #[error("malformed XML in part {part:?}: {message}")]
    MalformedXml {
        /// Part name whose XML failed to parse.
        part: String,
        /// Parser diagnostic.
        message: String,
    },
}

/// Whether a relationship points inside or outside the package.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TargetMode {
    /// The target is another part in this package.
    Internal,
    /// The target is an external resource (for example a hyperlink URL).
    External,
}

/// One `<Relationship>` from a `.rels` part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    /// The `Id` attribute, unique within its `.rels` part.
    pub id: String,
    /// The `Type` attribute (a schema URI).
    pub relationship_type: String,
    /// The raw `Target` attribute, unresolved. Use
    /// [`resolve_relationship_target`] to obtain a part name.
    pub target: String,
    /// Internal (default) or External per the `TargetMode` attribute.
    pub target_mode: TargetMode,
}

/// Where an entry's payload currently lives.
enum EntryPayload {
    /// Still in the source container; copied raw (compressed) on save.
    Untouched,
    /// Replaced in memory; re-deflated on save.
    Modified(Vec<u8>),
}

/// One ZIP entry of the source container.
///
/// Directory entries are kept so the saved container mirrors the source
/// entry list exactly; they are hidden from the part-level API.
struct PackageEntry {
    /// Zip entry name; directory entries keep their trailing `/`.
    name: String,
    /// Index of this entry in the source archive.
    source_index: usize,
    is_directory: bool,
    payload: EntryPayload,
}

/// A lossless in-memory OOXML package.
///
/// Invariant: `entries` lists every ZIP entry of the source container in its
/// original order, and saving writes them back in that order.
pub struct OpcPackage {
    source: Vec<u8>,
    entries: Vec<PackageEntry>,
}

impl fmt::Debug for OpcPackage {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("OpcPackage")
            .field("source_len", &self.source.len())
            .field("entries", &self.entries.len())
            .field("dirty", &self.dirty_part_names())
            .finish()
    }
}

impl OpcPackage {
    /// Load a package from a file on disk.
    pub fn from_path(path: impl AsRef<Path>) -> Result<Self, PackageError> {
        Self::from_bytes(std::fs::read(path)?)
    }

    /// Load a package from container bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self, PackageError> {
        let mut archive: zip::ZipArchive<Cursor<&[u8]>> =
            zip::ZipArchive::new(Cursor::new(bytes.as_slice()))?;
        let mut entries: Vec<PackageEntry> = Vec::with_capacity(archive.len());
        for source_index in 0..archive.len() {
            // Raw access skips decompression: payloads stay in `source` until
            // a caller actually reads or replaces them.
            let file = archive.by_index_raw(source_index)?;
            if file.last_modified().is_none() {
                return Err(PackageError::Container(
                    zip::result::ZipError::InvalidArchive(
                        "entry has an invalid DOS timestamp".into(),
                    ),
                ));
            }
            entries.push(PackageEntry {
                name: file.name()?.into_owned(),
                source_index,
                is_directory: file.is_dir(),
                payload: EntryPayload::Untouched,
            });
        }
        drop(archive);
        tracing::debug!(entry_count = entries.len(), "loaded OPC package");
        Ok(Self {
            source: bytes,
            entries,
        })
    }

    /// Names of every part (non-directory ZIP entry), in container order.
    pub fn part_names(&self) -> Vec<&str> {
        self.entries
            .iter()
            .filter(|entry| !entry.is_directory)
            .map(|entry| entry.name.as_str())
            .collect()
    }

    /// Whether a part with this exact name exists.
    pub fn has_part(&self, name: &str) -> bool {
        self.find_part(name).is_some()
    }

    /// The decompressed payload of a part.
    pub fn part_bytes(&self, name: &str) -> Result<Cow<'_, [u8]>, PackageError> {
        let entry: &PackageEntry = self
            .find_part(name)
            .ok_or_else(|| PackageError::MissingPart(name.to_string()))?;
        match &entry.payload {
            EntryPayload::Modified(bytes) => Ok(Cow::Borrowed(bytes.as_slice())),
            EntryPayload::Untouched => {
                let mut archive: zip::ZipArchive<Cursor<&[u8]>> =
                    zip::ZipArchive::new(Cursor::new(self.source.as_slice()))?;
                let mut file = archive.by_index(entry.source_index)?;
                let mut payload: Vec<u8> = Vec::with_capacity(file.size() as usize);
                file.read_to_end(&mut payload)?;
                Ok(Cow::Owned(payload))
            }
        }
    }

    /// Replace the payload of an existing part, marking it dirty.
    pub fn set_part_bytes(&mut self, name: &str, bytes: Vec<u8>) -> Result<(), PackageError> {
        let entry: &mut PackageEntry = self
            .entries
            .iter_mut()
            .find(|entry| !entry.is_directory && entry.name == name)
            .ok_or_else(|| PackageError::MissingPart(name.to_string()))?;
        tracing::debug!(
            part = name,
            byte_count = bytes.len(),
            "replacing part payload"
        );
        entry.payload = EntryPayload::Modified(bytes);
        Ok(())
    }

    /// Names of parts whose payload was replaced since load, in container order.
    pub fn dirty_part_names(&self) -> Vec<&str> {
        self.entries
            .iter()
            .filter(|entry| matches!(entry.payload, EntryPayload::Modified(_)))
            .map(|entry| entry.name.as_str())
            .collect()
    }

    /// Relationships of the package root (`None`) or of a part (`Some`).
    ///
    /// Returns an empty list when the corresponding `.rels` part is absent.
    pub fn relationships(
        &self,
        source_part: Option<&str>,
    ) -> Result<Vec<Relationship>, PackageError> {
        let rels_part: String = relationships_part_name(source_part);
        if !self.has_part(&rels_part) {
            return Ok(Vec::new());
        }
        let xml: Cow<'_, [u8]> = self.part_bytes(&rels_part)?;
        parse_relationships(&xml, &rels_part)
    }

    /// The declared content type of a part, from `[Content_Types].xml`.
    ///
    /// Override declarations win over extension defaults. Returns `None` when
    /// neither declares the part.
    pub fn content_type_of(&self, part_name: &str) -> Result<Option<String>, PackageError> {
        const CONTENT_TYPES_PART: &str = "[Content_Types].xml";
        if !self.has_part(CONTENT_TYPES_PART) {
            return Ok(None);
        }
        let xml: Cow<'_, [u8]> = self.part_bytes(CONTENT_TYPES_PART)?;
        let content_types: ContentTypes = parse_content_types(&xml, CONTENT_TYPES_PART)?;

        let absolute_name: String = format!("/{part_name}");
        if let Some(declared) = content_types
            .overrides
            .iter()
            .find(|(name, _)| *name == absolute_name)
        {
            return Ok(Some(declared.1.clone()));
        }
        let extension: Option<&str> = part_name.rsplit_once('.').map(|(_, ext)| ext);
        Ok(extension.and_then(|extension| {
            content_types
                .defaults
                .iter()
                // OPC compares extensions case-insensitively.
                .find(|(declared, _)| declared.eq_ignore_ascii_case(extension))
                .map(|(_, content_type)| content_type.clone())
        }))
    }

    /// Serialize the package back into container bytes.
    ///
    /// Untouched entries are copied raw (byte-for-byte compressed payload,
    /// preserving compression method and timestamps); dirty parts are
    /// re-deflated. Entry order is the source container's order.
    pub fn save_to_bytes(&self) -> Result<Vec<u8>, PackageError> {
        let mut source_archive: zip::ZipArchive<Cursor<&[u8]>> =
            zip::ZipArchive::new(Cursor::new(self.source.as_slice()))?;
        let mut writer: zip::ZipWriter<Cursor<Vec<u8>>> =
            zip::ZipWriter::new(Cursor::new(Vec::new()));
        let mut rewritten: usize = 0;
        for entry in &self.entries {
            match &entry.payload {
                EntryPayload::Untouched => {
                    let file = source_archive.by_index_raw(entry.source_index)?;
                    writer.raw_copy_file(file)?;
                }
                EntryPayload::Modified(bytes) => {
                    // Mirror the source entry's metadata so the rewrite is
                    // confined to the payload itself.
                    let source_file = source_archive.by_index_raw(entry.source_index)?;
                    let mut options: zip::write::SimpleFileOptions =
                        zip::write::SimpleFileOptions::default()
                            .compression_method(source_file.compression())
                            .last_modified_time(source_file.last_modified().ok_or_else(|| {
                                PackageError::Container(zip::result::ZipError::InvalidArchive(
                                    "entry has an invalid DOS timestamp".into(),
                                ))
                            })?);
                    if let Some(mode) = source_file.unix_mode() {
                        options = options.unix_permissions(mode);
                    }
                    drop(source_file);
                    writer.start_file(entry.name.as_str(), options)?;
                    writer.write_all(bytes)?;
                    rewritten += 1;
                }
            }
        }
        let cursor: Cursor<Vec<u8>> = writer.finish()?;
        tracing::debug!(
            entry_count = self.entries.len(),
            rewritten_count = rewritten,
            "saved OPC package"
        );
        Ok(cursor.into_inner())
    }

    fn find_part(&self, name: &str) -> Option<&PackageEntry> {
        self.entries
            .iter()
            .find(|entry| !entry.is_directory && entry.name == name)
    }
}

/// The `.rels` part name that holds relationships for `source_part`
/// (`None` addresses the package root).
pub fn relationships_part_name(source_part: Option<&str>) -> String {
    match source_part {
        None => "_rels/.rels".to_string(),
        Some(part) => match part.rsplit_once('/') {
            Some((directory, file_name)) => format!("{directory}/_rels/{file_name}.rels"),
            None => format!("_rels/{part}.rels"),
        },
    }
}

/// Resolve a relationship `Target` against its source part into a part name.
///
/// Handles targets relative to the source part's directory, `../` traversal,
/// and absolute (`/`-prefixed) targets. The result carries no leading slash,
/// matching ZIP entry names.
pub fn resolve_relationship_target(source_part: Option<&str>, target: &str) -> String {
    let mut segments: Vec<&str> = Vec::new();
    if !target.starts_with('/')
        && let Some((directory, _)) = source_part.and_then(|part| part.rsplit_once('/'))
    {
        segments.extend(directory.split('/'));
    }
    for segment in target.trim_start_matches('/').split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            other => segments.push(other),
        }
    }
    segments.join("/")
}

/// Parsed `[Content_Types].xml`: extension defaults and part-name overrides.
struct ContentTypes {
    /// `(extension, content type)` from `<Default>` elements.
    defaults: Vec<(String, String)>,
    /// `(absolute part name, content type)` from `<Override>` elements.
    overrides: Vec<(String, String)>,
}

fn parse_relationships(xml: &[u8], part: &str) -> Result<Vec<Relationship>, PackageError> {
    let mut reader: quick_xml::Reader<&[u8]> = quick_xml::Reader::from_reader(xml);
    let mut relationships: Vec<Relationship> = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if element.name().local_name().as_ref() == b"Relationship" =>
            {
                relationships.push(parse_relationship_element(&element, part)?);
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(PackageError::MalformedXml {
                    part: part.to_string(),
                    message: error.to_string(),
                });
            }
        }
    }
    Ok(relationships)
}

fn parse_relationship_element(
    element: &BytesStart<'_>,
    part: &str,
) -> Result<Relationship, PackageError> {
    let mut id: Option<String> = None;
    let mut relationship_type: Option<String> = None;
    let mut target: Option<String> = None;
    let mut target_mode: TargetMode = TargetMode::Internal;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| PackageError::MalformedXml {
            part: part.to_string(),
            message: error.to_string(),
        })?;
        let value: String = attribute
            .normalized_value(quick_xml::XmlVersion::Explicit1_0)
            .map_err(|error| PackageError::MalformedXml {
                part: part.to_string(),
                message: error.to_string(),
            })?
            .into_owned();
        match attribute.key.local_name().as_ref() {
            b"Id" => id = Some(value),
            b"Type" => relationship_type = Some(value),
            b"Target" => target = Some(value),
            b"TargetMode" if value == "External" => target_mode = TargetMode::External,
            _ => {}
        }
    }
    match (id, relationship_type, target) {
        (Some(id), Some(relationship_type), Some(target)) => Ok(Relationship {
            id,
            relationship_type,
            target,
            target_mode,
        }),
        _ => Err(PackageError::MalformedXml {
            part: part.to_string(),
            message: "Relationship element lacks Id, Type, or Target".to_string(),
        }),
    }
}

fn parse_content_types(xml: &[u8], part: &str) -> Result<ContentTypes, PackageError> {
    let mut reader: quick_xml::Reader<&[u8]> = quick_xml::Reader::from_reader(xml);
    let mut content_types: ContentTypes = ContentTypes {
        defaults: Vec::new(),
        overrides: Vec::new(),
    };
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                let is_default: bool = match element.name().local_name().as_ref() {
                    b"Default" => true,
                    b"Override" => false,
                    _ => continue,
                };
                let mut key: Option<String> = None;
                let mut content_type: Option<String> = None;
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| PackageError::MalformedXml {
                        part: part.to_string(),
                        message: error.to_string(),
                    })?;
                    let value: String = attribute
                        .normalized_value(quick_xml::XmlVersion::Explicit1_0)
                        .map_err(|error| PackageError::MalformedXml {
                            part: part.to_string(),
                            message: error.to_string(),
                        })?
                        .into_owned();
                    match attribute.key.local_name().as_ref() {
                        b"Extension" | b"PartName" => key = Some(value),
                        b"ContentType" => content_type = Some(value),
                        _ => {}
                    }
                }
                if let (Some(key), Some(content_type)) = (key, content_type) {
                    if is_default {
                        content_types.defaults.push((key, content_type));
                    } else {
                        content_types.overrides.push((key, content_type));
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => {
                return Err(PackageError::MalformedXml {
                    part: part.to_string(),
                    message: error.to_string(),
                });
            }
        }
    }
    Ok(content_types)
}
