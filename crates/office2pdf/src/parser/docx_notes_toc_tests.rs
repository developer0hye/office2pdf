use super::*;
use std::io::Cursor;

#[test]
fn footnote_references_keep_ids_and_position_without_a_reference_style() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:r><w:t>First</w:t><w:footnoteReference w:id="7"/><w:t> after.</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr></w:r><w:r><w:t>Second</w:t></w:r><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:footnoteReference w:id="29"/></w:r></w:p>
      <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t><w:footnoteReference w:id="11"/><w:t> tail.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
      <w:p><w:hyperlink w:anchor="target"><w:r><w:t>Linked</w:t><w:footnoteReference w:id="42"/><w:t> text.</w:t></w:r></w:hyperlink></w:p>
    </w:body></w:document>"#;
    let styles =
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
    let notes = r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:footnote w:id="29"><w:p><w:r><w:t>Second note.</w:t></w:r></w:p></w:footnote>
      <w:footnote w:id="7"><w:p><w:r><w:t>First note.</w:t></w:r></w:p></w:footnote>
      <w:footnote w:id="42"><w:p><w:r><w:t>Linked note.</w:t></w:r></w:p></w:footnote>
      <w:footnote w:id="11"><w:p><w:r><w:t>Cell note.</w:t></w:r></w:p></w:footnote>
    </w:footnotes>"#;
    let data = build_docx_with_notes_xml(xml, styles, notes);
    let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(flow) = &document.pages[0] else {
        panic!("expected flow")
    };
    let tokens = |paragraph: &Paragraph| -> Vec<String> {
        paragraph
            .runs
            .iter()
            .map(|run| {
                note_text(run).map_or_else(|| run.text.clone(), |text| format!("NOTE:{text}"))
            })
            .collect()
    };
    let Block::Paragraph(first) = &flow.content[0] else {
        panic!("first paragraph")
    };
    assert_eq!(tokens(first), ["First", "NOTE:First note.", " after."]);
    let Block::Paragraph(second) = &flow.content[1] else {
        panic!("second paragraph")
    };
    assert_eq!(tokens(second), ["Second", "NOTE:Second note."]);
    let Block::Table(table) = &flow.content[2] else {
        panic!("table")
    };
    let Block::Paragraph(cell) = &table.rows[0].cells[0].content[0] else {
        panic!("cell paragraph")
    };
    assert_eq!(tokens(cell), ["Cell", "NOTE:Cell note.", " tail."]);
    let Block::Paragraph(linked) = &flow.content[3] else {
        panic!("linked paragraph")
    };
    assert_eq!(tokens(linked), ["Linked", "NOTE:Linked note.", " text."]);
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn authored_unstyled_footnotes_remain_searchable_in_compiled_output() {
    let data: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/unstyled_footnotes.docx");
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let output = crate::render::typst_gen::generate_typst(&document).unwrap();
    let text_runs = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
    for expected in [
        "Invoice note remains searchable.",
        "Shipment note remains searchable.",
        "Closing paragraph remains searchable.",
    ] {
        assert!(
            text_runs.iter().any(|run| run.text.contains(expected)),
            "missing searchable text: {expected}"
        );
    }
}

#[test]
fn empty_normal_footnotes_keep_their_references_and_later_numbering() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Empty note</w:t><w:footnoteReference w:id="7"/><w:t> then populated note</w:t><w:footnoteReference w:id="29"/></w:r></w:p></w:body></w:document>"#;
    let styles =
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
    let notes = r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:t>Not a normal note.</w:t></w:r></w:p></w:footnote><w:footnote w:id="7"><w:p><w:r><w:footnoteRef/></w:r></w:p></w:footnote><w:footnote w:id="29"><w:p><w:r><w:footnoteRef/><w:t>Populated note.</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
    let data = build_docx_with_notes_xml(xml, styles, notes);
    let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let Page::Flow(flow) = &document.pages[0] else {
        panic!("expected flow")
    };
    let Block::Paragraph(paragraph) = &flow.content[0] else {
        panic!("paragraph")
    };
    let notes: Vec<&Run> = paragraph
        .runs
        .iter()
        .filter(|run| run.footnote.is_some())
        .collect();
    assert_eq!(notes.len(), 2, "empty normal notes still have a reference");
    assert_eq!(note_text(notes[0]).as_deref(), Some(""));
    assert_eq!(note_text(notes[1]).as_deref(), Some("Populated note."));
    #[cfg(not(target_arch = "wasm32"))]
    {
        let output = crate::render::typst_gen::generate_typst(&document).unwrap();
        let text_runs = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
        for number in ["1", "2"] {
            assert_eq!(
                text_runs.iter().filter(|run| run.text == number).count(),
                2,
                "both the body reference and entry retain number {number}"
            );
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn normal_footnote_markers_keep_reference_and_entry_run_formatting() {
    for (body_size, entry_size, family) in [
        (9, 14, "Arial"),
        (12, 10, "Times New Roman"),
        (18, 8, "Courier New"),
    ] {
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="{family}" w:hAnsi="{family}"/><w:sz w:val="{}"/></w:rPr><w:t>Body control</w:t><w:footnoteReference w:id="7"/></w:r></w:p></w:body></w:document>"#,
            body_size * 2
        );
        let styles =
            r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        let notes = format!(
            r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="7"><w:p><w:r><w:rPr><w:rFonts w:ascii="{family}" w:hAnsi="{family}"/><w:sz w:val="{}"/></w:rPr><w:footnoteRef/><w:t>Entry control</w:t></w:r></w:p></w:footnote></w:footnotes>"#,
            entry_size * 2
        );
        let data = build_docx_with_notes_xml(&xml, styles, &notes);
        let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
        let output = crate::render::typst_gen::generate_typst(&document).unwrap();
        let runs = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
        let body = runs
            .iter()
            .find(|run| run.text.contains("Body control"))
            .unwrap();
        let entry = runs
            .iter()
            .find(|run| run.text.contains("Entry control"))
            .unwrap();
        let markers: Vec<_> = runs.iter().filter(|run| run.text == "1").collect();
        assert_eq!(
            markers.len(),
            2,
            "both linked reference and entry markers survive"
        );
        let body_marker = markers
            .iter()
            .min_by(|a, b| a.baseline_pt.total_cmp(&b.baseline_pt))
            .unwrap();
        let entry_marker = markers
            .iter()
            .max_by(|a, b| a.baseline_pt.total_cmp(&b.baseline_pt))
            .unwrap();
        assert!(
            (body_marker.font_size_pt - f64::from(body_size)).abs() < 0.05,
            "body marker has source size: {body_marker:?}"
        );
        assert_eq!(body_marker.family, body.family);
        assert!(
            (body_marker.baseline_pt - body.baseline_pt).abs() < 0.05,
            "normal body marker stays on the baseline"
        );
        assert!(
            (entry_marker.font_size_pt - f64::from(entry_size)).abs() < 0.05,
            "entry marker has its own source size: {entry_marker:?}"
        );
        assert_eq!(entry_marker.family, entry.family);
        assert!(
            (entry_marker.baseline_pt - entry.baseline_pt).abs() < 0.05,
            "normal entry marker stays on the baseline"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn footnote_marker_character_styles_allow_explicit_baseline_overrides() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rStyle w:val="BodyReference"/><w:vertAlign w:val="baseline"/></w:rPr><w:t>Body control</w:t><w:footnoteReference w:id="7"/></w:r></w:p></w:body></w:document>"#;
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="character" w:styleId="BodyReference"><w:name w:val="Body Reference"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/><w:vertAlign w:val="superscript"/></w:rPr></w:style><w:style w:type="character" w:styleId="EntryReference"><w:name w:val="Entry Reference"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>"#;
    let notes = r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="7"><w:p><w:r><w:rPr><w:rStyle w:val="EntryReference"/><w:vertAlign w:val="baseline"/></w:rPr><w:footnoteRef/></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr><w:t>Entry control</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
    let data = build_docx_with_notes_xml(xml, styles, notes);
    let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let output = crate::render::typst_gen::generate_typst(&document).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
    let body = runs
        .iter()
        .find(|run| run.text.contains("Body control"))
        .unwrap();
    let entry = runs
        .iter()
        .find(|run| run.text.contains("Entry control"))
        .unwrap();
    let markers: Vec<_> = runs.iter().filter(|run| run.text == "1").collect();
    assert_eq!(markers.len(), 2);
    for (marker, control, size) in [(markers[0], body, 12.0), (markers[1], entry, 10.0)] {
        assert!(
            (marker.font_size_pt - size).abs() < 0.05,
            "explicit baseline retains the full marker size: {marker:?}"
        );
        assert!(
            (marker.baseline_pt - control.baseline_pt).abs() < 0.05,
            "explicit baseline clears the character style's superscript"
        );
        assert_eq!(marker.family, control.family);
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn absent_entry_reference_does_not_invent_a_searchable_note_label() {
    let data = include_bytes!("../../../../tests/fixtures/docx/docxide/unstyled_footnotes.docx");
    let mut archive = zip::ZipArchive::new(Cursor::new(data.as_slice())).unwrap();
    let document_xml = read_zip_text(&mut archive, "word/document.xml").unwrap();
    let styles_xml = read_zip_text(&mut archive, "word/styles.xml").unwrap();
    let notes_xml = read_zip_text(&mut archive, "word/footnotes.xml")
        .unwrap()
        .replacen("<w:footnoteRef/>", "", 1);
    let data = build_docx_with_notes_xml(&document_xml, &styles_xml, &notes_xml);
    let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let output = crate::render::typst_gen::generate_typst(&document).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
    assert_eq!(runs.iter().filter(|run| run.text == "1").count(), 2);
    assert_eq!(
        runs.iter().filter(|run| run.text == "2").count(),
        1,
        "the body reference survives, but the absent entry reference is not invented"
    );
    assert!(
        runs.iter()
            .any(|run| run.text.contains("Shipment note remains searchable."))
    );
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn superscript_footnote_markers_do_not_reenter_the_entry_show_rule() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr><w:t>Body control</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="7"/></w:r></w:p></w:body></w:document>"#;
    let styles =
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
    let notes = r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="7"><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/></w:rPr><w:t>Entry control</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>raised</w:t></w:r></w:p></w:footnote></w:footnotes>"#;
    let data = build_docx_with_notes_xml(xml, styles, notes);
    let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
    let output = crate::render::typst_gen::generate_typst(&document).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
    let body = runs
        .iter()
        .find(|run| run.text.contains("Body control"))
        .unwrap();
    let entry = runs
        .iter()
        .find(|run| run.text.contains("Entry control"))
        .unwrap();
    let markers: Vec<_> = runs.iter().filter(|run| run.text == "1").collect();
    assert_eq!(markers.len(), 2);
    for (marker, control) in [(markers[0], body), (markers[1], entry)] {
        assert!(marker.font_size_pt < control.font_size_pt * 0.8);
        assert!(marker.baseline_pt < control.baseline_pt - 1.0);
        assert_eq!(marker.family, control.family);
    }
    let Page::Flow(flow) = &document.pages[0] else {
        panic!("flow page")
    };
    let Block::Paragraph(paragraph) = &flow.content[0] else {
        panic!("paragraph")
    };
    let note = paragraph
        .runs
        .iter()
        .find_map(|run| run.footnote.as_deref())
        .unwrap();
    assert_eq!(
        note.runs
            .iter()
            .find(|run| run.text == "raised")
            .unwrap()
            .style
            .vertical_align,
        Some(VerticalTextAlign::Superscript)
    );
}

#[test]
fn authored_note_text_keeps_the_space_after_its_entry_reference() {
    let data = include_bytes!("../../../../tests/fixtures/docx/docxide/unstyled_footnotes.docx");
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let Page::Flow(flow) = &document.pages[0] else {
        panic!("flow")
    };
    let notes: Vec<_> = flow
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .flat_map(|paragraph| &paragraph.runs)
        .filter_map(|run| run.footnote.as_deref())
        .collect();
    assert_eq!(notes.len(), 2);
    assert_eq!(notes[0].runs[0].text, " Invoice note remains searchable.");
    assert_eq!(notes[1].runs[0].text, " Shipment note remains searchable.");
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn compiled_note_entry_retains_source_space_after_marker() {
    for spaces in [" ", "  "] {
        let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Body</w:t><w:footnoteReference w:id="7"/></w:r></w:p></w:body></w:document>"#;
        let styles =
            r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        let notes = format!(
            r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="7"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve">{spaces}Entry control</w:t></w:r></w:p></w:footnote></w:footnotes>"#
        );
        let data = build_docx_with_notes_xml(xml, styles, &notes);
        let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
        let output = crate::render::typst_gen::generate_typst(&document).unwrap();
        let runs = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
        let entry: String = runs
            .iter()
            .filter(|run| run.baseline_pt > 500.0)
            .map(|run| run.text.as_str())
            .collect();
        assert!(
            entry.contains(&format!("1{spaces}Entry control")),
            "source spacing is searchable: {entry:?}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn authored_script_markers_retain_each_reference_face_and_size() {
    let data =
        include_bytes!("../../../../tests/fixtures/docx/docxide/footnote_marker_scripts.docx");
    let (document, _) = DocxParser.parse(data, &ConvertOptions::default()).unwrap();
    let Page::Flow(flow) = &document.pages[0] else {
        panic!("flow page")
    };
    let references: Vec<_> = flow
        .content
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .flat_map(|paragraph| &paragraph.runs)
        .filter(|run| run.footnote.is_some())
        .collect();
    for (run, (family, size)) in references.iter().zip([
        ("Arial", 8.0),
        ("Times New Roman", 14.0),
        ("Courier New", 18.0),
    ]) {
        assert_eq!(run.style.font_family.as_deref(), Some(family));
        assert_eq!(run.style.font_size, Some(size));
        assert_eq!(
            run.style.vertical_align,
            Some(VerticalTextAlign::Superscript)
        );
    }
    assert_eq!(references.len(), 3);
    let output = crate::render::typst_gen::generate_typst(&document).unwrap();
    let runs = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
    let controls: Vec<_> = runs
        .iter()
        .filter(|run| run.text.contains("Body control"))
        .collect();
    assert_eq!(controls.len(), 3);
    for (index, control) in controls.iter().enumerate() {
        let number = (index + 1).to_string();
        let markers: Vec<_> = runs.iter().filter(|run| run.text == number).collect();
        assert_eq!(markers.len(), 2, "body and linked entry markers survive");
        for marker in markers {
            assert_eq!(marker.family, control.family);
            assert!(marker.font_size_pt < control.font_size_pt * 0.8);
            assert!(marker.font_size_pt > control.font_size_pt * 0.5);
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn reference_and_entry_vertical_alignment_are_independent() {
    for (body_superscript, entry_superscript) in [(false, true), (true, false)] {
        let body_align = if body_superscript {
            r#"<w:vertAlign w:val="superscript"/>"#
        } else {
            ""
        };
        let entry_align = if entry_superscript {
            r#"<w:vertAlign w:val="superscript"/>"#
        } else {
            ""
        };
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/></w:rPr><w:t>Body control</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/>{body_align}</w:rPr><w:footnoteReference w:id="7"/></w:r></w:p></w:body></w:document>"#
        );
        let styles =
            r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;
        let notes = format!(
            r#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="7"><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/>{entry_align}</w:rPr><w:footnoteRef/></w:r><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/></w:rPr><w:t>Entry control</w:t></w:r></w:p></w:footnote></w:footnotes>"#
        );
        let data = build_docx_with_notes_xml(&xml, styles, &notes);
        let (document, _) = DocxParser.parse(&data, &ConvertOptions::default()).unwrap();
        let output = crate::render::typst_gen::generate_typst(&document).unwrap();
        let runs = crate::render::pdf::compiled_text_runs(&output.source, 0).unwrap();
        let markers: Vec<_> = runs.iter().filter(|run| run.text == "1").collect();
        assert_eq!(markers.len(), 2);
        for (marker, control_text, superscript) in [
            (markers[0], "Body control", body_superscript),
            (markers[1], "Entry control", entry_superscript),
        ] {
            let control = runs
                .iter()
                .find(|run| run.text.contains(control_text))
                .unwrap();
            if superscript {
                assert!(marker.font_size_pt < control.font_size_pt * 0.8);
                assert!(marker.baseline_pt < control.baseline_pt - 1.0);
            } else {
                assert!((marker.font_size_pt - control.font_size_pt).abs() < 0.05);
                assert!((marker.baseline_pt - control.baseline_pt).abs() < 0.05);
            }
        }
    }
}

// ── Footnotes and endnotes ──────────────────────────────────────────

/// The note's text, joined across its runs.
///
/// A note carries styled runs rather than one string, so its content is read
/// the same way a paragraph's is.
fn note_text(run: &Run) -> Option<String> {
    Some(
        run.footnote
            .as_ref()?
            .iter()
            .map(|note_run| note_run.text.as_str())
            .collect(),
    )
}

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

    let para = match &flow.content[0] {
        Block::Paragraph(p) => p,
        _ => panic!("Expected paragraph"),
    };

    let note_run = para.runs.iter().find(|r| r.footnote.is_some());
    assert!(note_run.is_some(), "Expected a run with footnote content");
    assert_eq!(
        note_text(note_run.unwrap()).as_deref(),
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
    assert_eq!(note_text(note_runs[0]).as_deref(), Some("First note."));
    assert_eq!(note_text(note_runs[1]).as_deref(), Some("Second note."));
}

#[test]
fn test_endnote_parsed_as_footnote() {
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
        note_text(note_run.unwrap()).as_deref(),
        Some("This is an endnote.")
    );
}

fn build_docx_with_endnote(text: &str, endnote_id: usize, endnote_text: &str) -> Vec<u8> {
    use std::io::Write;
    use zip::ZipWriter;
    use zip::write::FileOptions;

    let buf = Vec::new();
    let mut zip = ZipWriter::new(Cursor::new(buf));
    let opts = FileOptions::default();

    zip.start_file("[Content_Types].xml", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>"#).unwrap();

    zip.start_file("_rels/.rels", opts).unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

    zip.start_file("word/_rels/document.xml.rels", opts)
        .unwrap();
    zip.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>"#).unwrap();

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

fn build_docx_with_toc(items: Vec<docx_rs::TableOfContentsItem>) -> Vec<u8> {
    let toc = items.into_iter().fold(
        docx_rs::TableOfContents::new()
            .heading_styles_range(1, 3)
            .alias("Table of contents"),
        |toc, item| toc.add_item(item),
    );

    let style1 = docx_rs::Style::new("Heading1", docx_rs::StyleType::Paragraph).name("Heading 1");
    let style2 = docx_rs::Style::new("Heading2", docx_rs::StyleType::Paragraph).name("Heading 2");

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
