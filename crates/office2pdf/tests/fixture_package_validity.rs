#![cfg(not(target_arch = "wasm32"))] // native-only: walks the repository's fixture tree
//! Package-validity gate for the OOXML fixtures this repository authors.
//!
//! Markup Compatibility and Extensibility (ECMA-376 Part 3) requires every
//! prefix listed in `mc:Ignorable` to be bound to a namespace in scope on that
//! element. Word for Mac enforces it: a package naming an unbound prefix fails
//! to open with no visible error, logging only `HrReadMetroFromPistm` and
//! `CmdExecStop {"CmdName":"FileOpen","CmdReturn":"Error"}`, so the AppleScript
//! exporters that produce native ground truth die on `active document`
//! instead (#1641). Our own parsers never read `mc:Ignorable`, so conversion
//! succeeds and nothing else in the suite notices the fixture is unopenable.

use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};

use quick_xml::Reader;
use quick_xml::events::Event;

const MARKUP_COMPATIBILITY_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// One `mc:Ignorable` token that no `xmlns:` declaration in scope binds.
#[derive(Debug, PartialEq, Eq)]
struct UndeclaredIgnorablePrefix {
    /// Qualified name of the element carrying the offending `mc:Ignorable`.
    element: String,
    /// The unbound prefix itself.
    prefix: String,
}

/// Report every `mc:Ignorable` token in `xml` that no in-scope `xmlns:`
/// declaration binds.
///
/// Only `mc:Ignorable` is inspected. `mc:ProcessContent` looks similar but
/// carries QNames of elements (`w:txbxContent`), not bare prefixes, so reading
/// its tokens as prefixes would invent failures.
fn undeclared_ignorable_prefixes(xml: &[u8]) -> Result<Vec<UndeclaredIgnorablePrefix>, String> {
    let mut reader: Reader<&[u8]> = Reader::from_reader(xml);
    // An empty root element still carries the declarations and the attribute,
    // so expand it into Start/End and keep one scope-stack rule for both.
    reader.config_mut().expand_empty_elements = true;

    let mut buffer: Vec<u8> = Vec::new();
    let mut scopes: Vec<HashMap<String, String>> = vec![HashMap::new()];
    let mut findings: Vec<UndeclaredIgnorablePrefix> = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) => {
                let mut scope: HashMap<String, String> = scopes.last().cloned().unwrap_or_default();

                // Bind first: a declaration on this element is in scope for its
                // own attributes.
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    let key: String = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    if let Some(prefix) = key.strip_prefix("xmlns:") {
                        let value: String =
                            String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                        scope.insert(prefix.to_owned(), value);
                    }
                }

                let element_name: String =
                    String::from_utf8_lossy(element.name().as_ref()).into_owned();
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    let key: String = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                    let Some((attribute_prefix, local_name)) = key.split_once(':') else {
                        continue;
                    };
                    if local_name != "Ignorable" {
                        continue;
                    }
                    // Follow the namespace, not the customary `mc` spelling.
                    if scope.get(attribute_prefix).map(String::as_str)
                        != Some(MARKUP_COMPATIBILITY_NAMESPACE)
                    {
                        continue;
                    }
                    let value: String =
                        String::from_utf8_lossy(attribute.value.as_ref()).into_owned();
                    for token in value.split_whitespace() {
                        if !scope.contains_key(token) {
                            findings.push(UndeclaredIgnorablePrefix {
                                element: element_name.clone(),
                                prefix: token.to_owned(),
                            });
                        }
                    }
                }

                scopes.push(scope);
            }
            Ok(Event::End(_)) => {
                // The root scope stays; a stray end tag must not pop it away.
                if scopes.len() > 1 {
                    scopes.pop();
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(error.to_string()),
        }
        buffer.clear();
    }

    Ok(findings)
}

// ---------------------------------------------------------------------------
// Detector behaviour
// ---------------------------------------------------------------------------

#[test]
fn accepts_a_package_part_that_binds_every_ignorable_prefix() {
    let xml: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
            mc:Ignorable="w14 wp14"><w:body/></w:document>"#;

    assert_eq!(undeclared_ignorable_prefixes(xml), Ok(Vec::new()));
}

#[test]
fn reports_an_unbound_prefix_on_the_root_element() {
    let xml: &[u8] =
        br#"<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            mc:Ignorable="w14 wp14"><w:body/></w:document>"#;

    assert_eq!(
        undeclared_ignorable_prefixes(xml),
        Ok(vec![UndeclaredIgnorablePrefix {
            element: "w:document".to_owned(),
            prefix: "wp14".to_owned(),
        }])
    );
}

#[test]
fn reports_an_unbound_prefix_on_an_empty_root_element() {
    // `w:hdr` parts are often written as a single self-closing tag.
    let xml: &[u8] =
        br#"<w:hdr xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            mc:Ignorable="w14"/>"#;

    assert_eq!(
        undeclared_ignorable_prefixes(xml),
        Ok(vec![UndeclaredIgnorablePrefix {
            element: "w:hdr".to_owned(),
            prefix: "w14".to_owned(),
        }])
    );
}

#[test]
fn reports_an_unbound_prefix_on_a_nested_element() {
    let xml: &[u8] =
        br#"<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><mc:AlternateContent mc:Ignorable="v"/></w:body>
</w:document>"#;

    assert_eq!(
        undeclared_ignorable_prefixes(xml),
        Ok(vec![UndeclaredIgnorablePrefix {
            element: "mc:AlternateContent".to_owned(),
            prefix: "v".to_owned(),
        }])
    );
}

#[test]
fn accepts_a_prefix_bound_by_an_ancestor_element() {
    let xml: &[u8] =
        br#"<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body><w:p mc:Ignorable="v"/></w:body>
</w:document>"#;

    assert_eq!(undeclared_ignorable_prefixes(xml), Ok(Vec::new()));
}

#[test]
fn stops_honouring_a_binding_once_its_element_closes() {
    let xml: &[u8] =
        br#"<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body xmlns:v="urn:schemas-microsoft-com:vml"><w:p mc:Ignorable="v"/></w:body>
  <w:p mc:Ignorable="v"/>
</w:document>"#;

    assert_eq!(
        undeclared_ignorable_prefixes(xml),
        Ok(vec![UndeclaredIgnorablePrefix {
            element: "w:p".to_owned(),
            prefix: "v".to_owned(),
        }])
    );
}

#[test]
fn follows_the_compatibility_namespace_rather_than_the_mc_spelling() {
    // `Ignorable` under a differently spelled prefix still counts ...
    let renamed: &[u8] =
        br#"<w:document xmlns:compat="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            compat:Ignorable="w14"/>"#;
    assert_eq!(
        undeclared_ignorable_prefixes(renamed),
        Ok(vec![UndeclaredIgnorablePrefix {
            element: "w:document".to_owned(),
            prefix: "w14".to_owned(),
        }])
    );

    // ... while an `mc` bound elsewhere is a different attribute entirely.
    let foreign: &[u8] = br#"<w:document xmlns:mc="urn:example:not-markup-compatibility"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            mc:Ignorable="w14"/>"#;
    assert_eq!(undeclared_ignorable_prefixes(foreign), Ok(Vec::new()));
}

#[test]
fn leaves_process_content_qnames_alone() {
    // `mc:ProcessContent` lists element QNames, not prefixes; its `w:` token
    // must not be read as an undeclared prefix named `w:txbxContent`.
    let xml: &[u8] = br#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            mc:ProcessContent="w:txbxContent"/>"#;

    assert_eq!(undeclared_ignorable_prefixes(xml), Ok(Vec::new()));
}

// ---------------------------------------------------------------------------
// Fixture corpus
// ---------------------------------------------------------------------------

fn tests_directory() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../tests")
}

/// Bulk third-party corpora and confidential material. `poi/` and
/// `libreoffice/` are other projects' regression inputs — several are truncated
/// or encrypted on purpose — and `classified_fixtures/` is untracked private
/// documents. Neither is ours to keep openable.
const SKIPPED_DIRECTORIES: [&str; 3] = ["poi", "libreoffice", "classified_fixtures"];

const OOXML_EXTENSIONS: [&str; 9] = [
    "docx", "docm", "dotx", "pptx", "pptm", "potx", "xlsx", "xlsm", "xltx",
];

fn collect_authored_packages(directory: &Path, packages: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(directory) else {
        return;
    };
    for entry in entries.flatten() {
        let path: PathBuf = entry.path();
        let name: String = entry.file_name().to_string_lossy().into_owned();
        if path.is_dir() {
            if !SKIPPED_DIRECTORIES.contains(&name.as_str()) {
                collect_authored_packages(&path, packages);
            }
            continue;
        }
        let extension: String = path
            .extension()
            .map(|value| value.to_string_lossy().to_lowercase())
            .unwrap_or_default();
        if OOXML_EXTENSIONS.contains(&extension.as_str()) {
            packages.push(path);
        }
    }
}

/// Every prefix named in `mc:Ignorable` must be declared in scope, in every
/// tracked package this repository authors — otherwise Word refuses to open the
/// fixture and no native ground truth can be exported from it (#1641).
#[test]
fn authored_fixtures_declare_every_ignorable_prefix() {
    let mut packages: Vec<PathBuf> = Vec::new();
    collect_authored_packages(&tests_directory(), &mut packages);
    packages.sort();

    let mut failures: Vec<String> = Vec::new();
    let mut parts_with_ignorable: usize = 0;

    for package in &packages {
        let Ok(file) = File::open(package) else {
            continue;
        };
        // A package that is not a ZIP container cannot be inspected here; the
        // parser-tolerance fixtures are deliberately truncated or encrypted.
        let Ok(mut archive) = zip::ZipArchive::new(file) else {
            continue;
        };
        let part_names: Vec<String> = archive.file_names().map(str::to_owned).collect();
        for part_name in part_names {
            if !part_name.ends_with(".xml") && !part_name.ends_with(".rels") {
                continue;
            }
            let Ok(mut part) = archive.by_name(&part_name) else {
                continue;
            };
            let mut xml: Vec<u8> = Vec::new();
            if part.read_to_end(&mut xml).is_err() {
                continue;
            }
            if !xml.windows(10).any(|window| window == b"Ignorable=") {
                continue;
            }
            parts_with_ignorable += 1;
            match undeclared_ignorable_prefixes(&xml) {
                Ok(findings) => {
                    for finding in findings {
                        failures.push(format!(
                            "{}: {} <{}> names undeclared prefix '{}' in mc:Ignorable",
                            package.display(),
                            part_name,
                            finding.element,
                            finding.prefix
                        ));
                    }
                }
                // Malformed XML is a different defect with its own fixtures.
                Err(_) => continue,
            }
        }
    }

    // Guard against a silently empty sweep: the corpus is large and most of its
    // Word and PowerPoint parts carry `mc:Ignorable`.
    assert!(
        packages.len() > 50,
        "expected to sweep the fixture corpus, found only {} packages under {}",
        packages.len(),
        tests_directory().display()
    );
    assert!(
        parts_with_ignorable > 20,
        "expected the sweep to read mc:Ignorable parts, found {parts_with_ignorable}"
    );
    assert!(
        failures.is_empty(),
        "packages Word refuses to open:\n{}",
        failures.join("\n")
    );
}
