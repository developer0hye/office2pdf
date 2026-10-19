//! Lossless round-trip tests for the OPC package model (issue #746).
//!
//! Every assertion here reads the saved container back with the `zip` crate
//! directly, so the crate under test never verifies itself.

use std::io::{Cursor, Read, Write};
use std::path::PathBuf;

use ooxml_package::{OpcPackage, PackageError, TargetMode, resolve_relationship_target};

/// OPC is format-agnostic: the round-trip guarantees must hold for all three
/// OOXML flavors, not just the PPTX milestone target.
const ROUND_TRIP_FIXTURES: [&str; 4] = [
    "pptx/minimal.pptx",
    "pptx/WithMaster.pptx",
    "docx/unit_test_headers.docx",
    "xlsx/temperature.xlsx",
];

fn fixture_path(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures")
        .join(relative)
}

fn fixture_bytes(relative: &str) -> Vec<u8> {
    std::fs::read(fixture_path(relative)).expect("fixture should be readable")
}

/// `(name, decompressed payload)` for every non-directory entry, in order.
fn zip_entries(bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("container should open");
    let mut entries: Vec<(String, Vec<u8>)> = Vec::new();
    for index in 0..archive.len() {
        let mut file = archive.by_index(index).expect("entry should open");
        if file.is_dir() {
            continue;
        }
        let mut payload: Vec<u8> = Vec::new();
        file.read_to_end(&mut payload)
            .expect("entry should decompress");
        entries.push((
            file.name().expect("entry name should decode").into_owned(),
            payload,
        ));
    }
    entries
}

/// `(name, raw compressed payload)` for every non-directory entry, in order.
fn raw_zip_entries(bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("container should open");
    let mut entries: Vec<(String, Vec<u8>)> = Vec::new();
    for index in 0..archive.len() {
        let mut file = archive.by_index_raw(index).expect("raw entry should open");
        if file.is_dir() {
            continue;
        }
        let mut payload: Vec<u8> = Vec::new();
        file.read_to_end(&mut payload)
            .expect("raw entry should read");
        entries.push((
            file.name().expect("entry name should decode").into_owned(),
            payload,
        ));
    }
    entries
}

/// Rebuild `original` with extra entries appended, using the `zip` crate only.
fn with_extra_entries(original: &[u8], extras: &[(&str, &[u8])]) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(original)).expect("container should open");
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let file = archive.by_index_raw(index).expect("raw entry should open");
        writer.raw_copy_file(file).expect("raw copy should succeed");
    }
    for (name, payload) in extras {
        writer
            .start_file(*name, zip::write::SimpleFileOptions::default())
            .expect("extra entry should start");
        writer.write_all(payload).expect("extra entry should write");
    }
    writer
        .finish()
        .expect("container should finish")
        .into_inner()
}

fn patch_entry_dos_date(bytes: &mut [u8], name: &str, date: u16) {
    let mut patched_headers: usize = 0;
    for offset in 0..bytes.len().saturating_sub(4) {
        let signature: [u8; 4] = bytes[offset..offset + 4]
            .try_into()
            .expect("signature slice should have four bytes");
        let (name_length_offset, name_offset, date_offset) = match signature {
            [0x50, 0x4b, 0x03, 0x04] => (26, 30, 12),
            [0x50, 0x4b, 0x01, 0x02] => (28, 46, 14),
            _ => continue,
        };
        let name_length = u16::from_le_bytes([
            bytes[offset + name_length_offset],
            bytes[offset + name_length_offset + 1],
        ]) as usize;
        let name_start = offset + name_offset;
        let name_end = name_start + name_length;
        if bytes.get(name_start..name_end) == Some(name.as_bytes()) {
            bytes[offset + date_offset..offset + date_offset + 2]
                .copy_from_slice(&date.to_le_bytes());
            patched_headers += 1;
        }
    }
    assert_eq!(patched_headers, 2, "local and central headers should match");
}

fn entry_header_timestamps(bytes: &[u8], name: &str) -> Vec<(u16, u16)> {
    let mut timestamps: Vec<(u16, u16)> = Vec::new();
    for offset in 0..bytes.len().saturating_sub(4) {
        let signature: [u8; 4] = bytes[offset..offset + 4]
            .try_into()
            .expect("signature slice should have four bytes");
        let (name_length_offset, name_offset, time_offset, date_offset) = match signature {
            [0x50, 0x4b, 0x03, 0x04] => (26, 30, 10, 12),
            [0x50, 0x4b, 0x01, 0x02] => (28, 46, 12, 14),
            _ => continue,
        };
        let name_length = u16::from_le_bytes([
            bytes[offset + name_length_offset],
            bytes[offset + name_length_offset + 1],
        ]) as usize;
        let name_start = offset + name_offset;
        let name_end = name_start + name_length;
        if bytes.get(name_start..name_end) == Some(name.as_bytes()) {
            let time =
                u16::from_le_bytes([bytes[offset + time_offset], bytes[offset + time_offset + 1]]);
            let date =
                u16::from_le_bytes([bytes[offset + date_offset], bytes[offset + date_offset + 1]]);
            timestamps.push((time, date));
        }
    }
    timestamps
}

#[test]
fn no_op_round_trip_preserves_every_part_payload_and_order() {
    for fixture in ROUND_TRIP_FIXTURES {
        let original: Vec<u8> = fixture_bytes(fixture);
        let package = OpcPackage::from_bytes(original.clone())
            .unwrap_or_else(|error| panic!("{fixture} should load: {error}"));
        let saved: Vec<u8> = package
            .save_to_bytes()
            .unwrap_or_else(|error| panic!("{fixture} should save: {error}"));
        assert_eq!(
            zip_entries(&original),
            zip_entries(&saved),
            "{fixture}: part names, order, or payloads changed across a no-op round trip"
        );
    }
}

#[test]
fn package_rejects_invalid_dos_date_before_a_rewrite_can_normalize_it() {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    writer
        .start_file("ppt/document.xml", zip::write::SimpleFileOptions::default())
        .expect("part should start");
    writer.write_all(b"<document/>").expect("part should write");
    let mut original = writer.finish().expect("ZIP should finish").into_inner();
    patch_entry_dos_date(&mut original, "ppt/document.xml", 0);

    let error = match OpcPackage::from_bytes(original) {
        Err(error) => error,
        Ok(_) => panic!("invalid DOS date should be rejected"),
    };
    assert!(
        matches!(&error, PackageError::Container(zip::result::ZipError::InvalidArchive(message)) if message.contains("invalid DOS timestamp")),
        "unexpected invalid DOS date error: {error}"
    );
}

#[test]
fn modified_part_keeps_untouched_part_dos_timestamps() {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let timestamp =
        zip::DateTime::from_date_and_time(2022, 4, 3, 2, 1, 0).expect("timestamp should be valid");
    let options = zip::write::SimpleFileOptions::default().last_modified_time(timestamp);
    writer
        .start_file("ppt/document.xml", options)
        .expect("first part should start");
    writer
        .write_all(b"<document/>")
        .expect("first part should write");
    writer
        .start_file("ppt/presentation.xml", options)
        .expect("second part should start");
    writer
        .write_all(b"<presentation/>")
        .expect("second part should write");
    let original = writer.finish().expect("ZIP should finish").into_inner();

    let mut package = OpcPackage::from_bytes(original.clone()).expect("package should load");
    package
        .set_part_bytes("ppt/document.xml", b"<document changed=\"true\"/>".to_vec())
        .expect("first part should update");
    let saved = package.save_to_bytes().expect("package should save");

    assert_eq!(
        entry_header_timestamps(&original, "ppt/presentation.xml"),
        entry_header_timestamps(&saved, "ppt/presentation.xml"),
        "both raw-copied timestamp fields should remain byte-for-byte unchanged"
    );
}

#[test]
fn untouched_part_payloads_are_copied_byte_for_byte_compressed() {
    for fixture in ROUND_TRIP_FIXTURES {
        let original: Vec<u8> = fixture_bytes(fixture);
        let package = OpcPackage::from_bytes(original.clone()).expect("fixture should load");
        let saved: Vec<u8> = package.save_to_bytes().expect("package should save");
        assert_eq!(
            raw_zip_entries(&original),
            raw_zip_entries(&saved),
            "{fixture}: untouched entries should keep their exact compressed payload"
        );
    }
}

#[test]
fn part_bytes_returns_the_decompressed_payload() {
    let original: Vec<u8> = fixture_bytes("pptx/WithMaster.pptx");
    let package = OpcPackage::from_bytes(original.clone()).expect("fixture should load");
    let expected: Vec<(String, Vec<u8>)> = zip_entries(&original);
    for (name, payload) in &expected {
        let actual = package.part_bytes(name).expect("part should be readable");
        assert_eq!(
            actual.as_ref(),
            payload.as_slice(),
            "part {name} should decompress to the original payload"
        );
    }
    assert_eq!(
        package.part_names(),
        expected
            .iter()
            .map(|(name, _)| name.as_str())
            .collect::<Vec<&str>>(),
        "part_names should list every part in container order"
    );
}

#[test]
fn unknown_parts_survive_a_round_trip() {
    let opaque_payload: Vec<u8> = (0u8..=255).cycle().take(1024).collect();
    let extras: [(&str, &[u8]); 2] = [
        (
            "customXml/unknown-tool-data.xml",
            b"<unknownTool><keep everything=\"intact\"/></unknownTool>".as_slice(),
        ),
        ("ppt/media/opaque-blob.bin", opaque_payload.as_slice()),
    ];
    let with_unknown: Vec<u8> = with_extra_entries(&fixture_bytes("pptx/WithMaster.pptx"), &extras);

    let package = OpcPackage::from_bytes(with_unknown.clone()).expect("package should load");
    for (name, payload) in &extras {
        assert!(
            package.has_part(name),
            "unknown part {name} should be visible"
        );
        assert_eq!(
            package
                .part_bytes(name)
                .expect("unknown part should read")
                .as_ref(),
            *payload,
            "unknown part {name} should keep its payload in memory"
        );
    }

    let saved: Vec<u8> = package.save_to_bytes().expect("package should save");
    assert_eq!(
        zip_entries(&with_unknown),
        zip_entries(&saved),
        "unknown parts must survive a no-op round trip untouched"
    );
}

#[test]
fn modifying_one_part_rewrites_only_that_entry() {
    let original: Vec<u8> = fixture_bytes("pptx/WithMaster.pptx");
    let target_part: &str = "ppt/slides/slide2.xml";

    let mut package = OpcPackage::from_bytes(original.clone()).expect("fixture should load");
    let original_slide: Vec<u8> = package
        .part_bytes(target_part)
        .expect("slide should be readable")
        .into_owned();
    let modified_slide: Vec<u8> = String::from_utf8(original_slide.clone())
        .expect("slide XML should be UTF-8")
        .replace("subtitle", "caption")
        .into_bytes();
    assert_ne!(
        original_slide, modified_slide,
        "the edit should change the slide"
    );

    package
        .set_part_bytes(target_part, modified_slide.clone())
        .expect("existing part should accept new bytes");
    assert_eq!(
        package.dirty_part_names(),
        vec![target_part],
        "exactly the edited part should be dirty"
    );

    let saved: Vec<u8> = package.save_to_bytes().expect("package should save");
    let original_raw: Vec<(String, Vec<u8>)> = raw_zip_entries(&original);
    let saved_raw: Vec<(String, Vec<u8>)> = raw_zip_entries(&saved);
    assert_eq!(
        original_raw
            .iter()
            .map(|(name, _)| name.as_str())
            .collect::<Vec<&str>>(),
        saved_raw
            .iter()
            .map(|(name, _)| name.as_str())
            .collect::<Vec<&str>>(),
        "entry names and order should be preserved"
    );
    for ((name, original_payload), (_, saved_payload)) in original_raw.iter().zip(saved_raw.iter())
    {
        if name == target_part {
            continue;
        }
        assert_eq!(
            original_payload, saved_payload,
            "untouched entry {name} should keep its exact compressed payload"
        );
    }
    assert_eq!(
        zip_entries(&saved)
            .into_iter()
            .find(|(name, _)| name == target_part)
            .expect("edited part should exist")
            .1,
        modified_slide,
        "the edited part should carry the replacement payload"
    );
}

#[test]
fn relationships_resolve_to_existing_parts_before_and_after_round_trip() {
    for fixture in ROUND_TRIP_FIXTURES {
        let package = OpcPackage::from_bytes(fixture_bytes(fixture)).expect("fixture should load");
        assert_relationships_resolve(&package, fixture);

        let reloaded =
            OpcPackage::from_bytes(package.save_to_bytes().expect("package should save"))
                .expect("saved container should reload");
        assert_relationships_resolve(&reloaded, fixture);
    }
}

fn assert_relationships_resolve(package: &OpcPackage, fixture: &str) {
    let root_relationships: Vec<ooxml_package::Relationship> = package
        .relationships(None)
        .expect("root relationships should parse");
    let office_document = root_relationships
        .iter()
        .find(|relationship| relationship.relationship_type.ends_with("/officeDocument"))
        .unwrap_or_else(|| panic!("{fixture} should declare an officeDocument relationship"));
    assert!(
        package.has_part(&resolve_relationship_target(None, &office_document.target)),
        "{fixture}: officeDocument target should exist"
    );

    let part_names: Vec<String> = package
        .part_names()
        .into_iter()
        .map(str::to_string)
        .collect();
    for part_name in &part_names {
        let relationships = package
            .relationships(Some(part_name))
            .expect("part relationships should parse");
        for relationship in relationships {
            if relationship.target_mode == TargetMode::External {
                continue;
            }
            let resolved: String =
                resolve_relationship_target(Some(part_name), &relationship.target);
            assert!(
                package.has_part(&resolved),
                "{fixture}: relationship {id} of {part_name} targets missing part {resolved}",
                id = relationship.id,
            );
        }
    }
}

#[test]
fn content_types_cover_every_part() {
    for fixture in ROUND_TRIP_FIXTURES {
        let package = OpcPackage::from_bytes(fixture_bytes(fixture)).expect("fixture should load");
        for part_name in package.part_names() {
            if part_name == "[Content_Types].xml" {
                continue;
            }
            let content_type: Option<String> = package
                .content_type_of(part_name)
                .expect("content types should parse");
            assert!(
                content_type.is_some_and(|declared| !declared.is_empty()),
                "{fixture}: part {part_name} should have a declared content type"
            );
        }
    }
}

#[test]
fn from_bytes_rejects_bytes_that_are_not_a_zip_container() {
    let error: PackageError =
        OpcPackage::from_bytes(b"this is not a zip container".to_vec()).unwrap_err();
    assert!(
        matches!(error, PackageError::Container(_)),
        "expected a container error, got: {error}"
    );
}

#[test]
fn missing_parts_are_reported_by_name() {
    let mut package =
        OpcPackage::from_bytes(fixture_bytes("pptx/minimal.pptx")).expect("fixture should load");
    let read_error: PackageError = package.part_bytes("ppt/absent.xml").unwrap_err();
    assert!(
        matches!(&read_error, PackageError::MissingPart(name) if name == "ppt/absent.xml"),
        "expected MissingPart, got: {read_error}"
    );
    let write_error: PackageError = package
        .set_part_bytes("ppt/absent.xml", vec![1, 2, 3])
        .unwrap_err();
    assert!(
        matches!(&write_error, PackageError::MissingPart(name) if name == "ppt/absent.xml"),
        "expected MissingPart, got: {write_error}"
    );
}

#[test]
fn relationship_targets_resolve_relative_to_their_source_part() {
    assert_eq!(
        resolve_relationship_target(None, "ppt/presentation.xml"),
        "ppt/presentation.xml"
    );
    assert_eq!(
        resolve_relationship_target(Some("ppt/presentation.xml"), "slides/slide1.xml"),
        "ppt/slides/slide1.xml"
    );
    assert_eq!(
        resolve_relationship_target(
            Some("ppt/slideLayouts/slideLayout1.xml"),
            "../slideMasters/slideMaster1.xml"
        ),
        "ppt/slideMasters/slideMaster1.xml"
    );
    assert_eq!(
        resolve_relationship_target(Some("ppt/presentation.xml"), "/docProps/core.xml"),
        "docProps/core.xml"
    );
}
