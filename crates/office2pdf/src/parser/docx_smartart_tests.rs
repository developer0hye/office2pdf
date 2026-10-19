use super::*;
use std::io::{Read, Write};

const SMARTART_REPORT: &[u8] =
    include_bytes!("../../../../tests/fixtures/docx/docxide/smartart_cached_drawing.docx");

fn drawing_groups(blocks: &[Block]) -> Vec<&crate::ir::DrawingGroup> {
    let mut groups = Vec::new();
    for block in blocks {
        match block {
            Block::DrawingGroup(group) => groups.push(group.as_ref()),
            Block::Paragraph(paragraph) => {
                for run in &paragraph.runs {
                    if let Some(frame) = &run.inline_box {
                        groups.extend(drawing_groups(&frame.content));
                    }
                }
            }
            Block::FloatingTextBox(frame) => groups.extend(drawing_groups(&frame.content)),
            _ => {}
        }
    }
    groups
}

fn rewrite_part(package: &[u8], part: &str, transform: impl Fn(&str) -> String) -> Vec<u8> {
    let mut source = zip::ZipArchive::new(Cursor::new(package)).unwrap();
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..source.len() {
        let mut entry = source.by_index(index).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        let name = entry.name().expect("entry name should decode").into_owned();
        if name == part {
            bytes = transform(std::str::from_utf8(&bytes).unwrap()).into_bytes();
        }
        output
            .start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        output.write_all(&bytes).unwrap();
    }
    output.finish().unwrap().into_inner()
}

#[test]
fn docx_cached_smartart_inline_frame_stays_in_its_text_paragraph() {
    let package = rewrite_part(SMARTART_REPORT, "word/document.xml", |xml| {
        let start = xml.find("<w:p ").unwrap();
        let content = xml[start..].find('>').unwrap() + start + 1;
        let end = xml[content..].find("</w:p>").unwrap() + content;
        let mut modified = xml.to_string();
        modified.insert_str(end, "<w:r><w:t> End</w:t></w:r>");
        modified.insert_str(content, "<w:r><w:t>Start </w:t></w:r>");
        modified
    });
    let (document, _) = DocxParser
        .parse(&package, &ConvertOptions::default())
        .unwrap();
    let paragraph = all_blocks(&document)
        .iter()
        .filter_map(|block| match block {
            Block::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .find(|paragraph| paragraph.runs.iter().any(|run| run.text.contains("Start")))
        .expect("anchor paragraph");
    assert!(paragraph.runs.iter().any(|run| run.text.contains("End")));
    assert!(paragraph.runs.iter().any(|run| run.inline_box.is_some()));
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_cached_smartart_keeps_searchable_diagram_text() {
    let result = crate::convert_bytes(
        SMARTART_REPORT,
        crate::Format::Docx,
        &ConvertOptions::default(),
    )
    .expect("cached SmartArt must render as document content");
    let text = pdf_extract::extract_text_from_mem(&result.pdf).unwrap();
    for label in ["Foo", "Shmoo", "Lorem", "Ipsum", "Dolor", "Amet"] {
        assert!(text.contains(label), "missing diagram label {label}");
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_cached_smartart_follows_relationships_after_parts_and_prefixes_change() {
    let mut source = zip::ZipArchive::new(Cursor::new(SMARTART_REPORT)).unwrap();
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..source.len() {
        let mut entry = source.by_index(index).unwrap();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        let name = entry
            .name()
            .expect("entry name should decode")
            .into_owned()
            .replace("drawing1.xml", "diagram-cache-alpha.xml")
            .replace("data1.xml", "diagram-model-alpha.xml");
        if name.ends_with(".xml") || name.ends_with(".rels") {
            let xml = String::from_utf8(bytes)
                .unwrap()
                .replace("drawing1.xml", "diagram-cache-alpha.xml")
                .replace("data1.xml", "diagram-model-alpha.xml")
                .replace("xmlns:w=", "xmlns:word=")
                .replace("<w:", "<word:")
                .replace("</w:", "</word:")
                .replace(" w:", " word:")
                .replace("xmlns:dsp=", "xmlns:cached=")
                .replace("<dsp:", "<cached:")
                .replace("</dsp:", "</cached:");
            bytes = xml.into_bytes();
        }
        output
            .start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        output.write_all(&bytes).unwrap();
    }
    let package = output.finish().unwrap().into_inner();
    let result =
        crate::convert_bytes(&package, crate::Format::Docx, &ConvertOptions::default()).unwrap();
    let text = pdf_extract::extract_text_from_mem(&result.pdf).unwrap();
    for label in ["Foo", "Shmoo", "Lorem", "Ipsum", "Dolor", "Amet"] {
        assert!(text.contains(label), "missing diagram label {label}");
    }
}

#[test]
fn docx_cached_smartart_preserves_drawing_extent_and_vector_content() {
    let (document, _warnings) = DocxParser
        .parse(SMARTART_REPORT, &ConvertOptions::default())
        .unwrap();
    let drawings = drawing_groups(all_blocks(&document));
    assert_eq!(drawings.len(), 2);
    for drawing in drawings {
        assert_eq!((drawing.width, drawing.height), (432.0, 252.0));
        assert!(drawing.elements.iter().any(|element| matches!(
            element.kind,
            crate::ir::FixedElementKind::Shape(_) | crate::ir::FixedElementKind::TextBox(_)
        )));
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_cached_smartart_ignores_diagrams_in_unrendered_story_branches() {
    for wrapper in ["deleted", "moved", "fallback"] {
        let package = rewrite_part(SMARTART_REPORT, "word/document.xml", |xml| {
            let start = xml.find("<w:drawing>").unwrap();
            let end = xml[start..].find("</w:drawing>").unwrap() + start + "</w:drawing>".len();
            let drawing = &xml[start..end];
            let run = format!("<w:r>{drawing}</w:r>");
            let hidden = match wrapper {
                "deleted" => format!("<w:del w:id=\"42\" w:author=\"Synthetic\">{run}</w:del>"),
                "moved" => {
                    format!("<w:moveFrom w:id=\"43\" w:author=\"Synthetic\">{run}</w:moveFrom>")
                }
                "fallback" => format!(
                    "<mc:AlternateContent><mc:Choice Requires=\"w14\"/><mc:Fallback>{run}</mc:Fallback></mc:AlternateContent>"
                ),
                _ => unreachable!(),
            };
            xml.replacen("<w:body>", &format!("<w:body><w:p>{hidden}</w:p>"), 1)
        });
        let output =
            crate::convert_bytes(&package, crate::Format::Docx, &ConvertOptions::default())
                .unwrap();
        let text = pdf_extract::extract_text_from_mem(&output.pdf).unwrap();
        for label in ["Foo", "Shmoo", "Lorem", "Ipsum", "Dolor", "Amet"] {
            assert!(
                text.contains(label),
                "{wrapper} displaced visible diagram label {label}"
            );
        }
        assert_eq!(
            text.matches("Foo").count(),
            1,
            "{wrapper} duplicated the first diagram"
        );
    }
}

#[test]
fn docx_cached_smartart_keeps_a_neighboring_picture_source_crop() {
    let picture_package: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/floating_picture_stretch.docx");
    let mut pictures = zip::ZipArchive::new(Cursor::new(picture_package)).unwrap();
    let mut picture_xml = String::new();
    pictures
        .by_name("word/document.xml")
        .unwrap()
        .read_to_string(&mut picture_xml)
        .unwrap();
    let drawing_start = picture_xml.match_indices("<w:drawing>").nth(2).unwrap().0;
    let drawing_end = picture_xml[drawing_start..].find("</w:drawing>").unwrap()
        + drawing_start
        + "</w:drawing>".len();
    let drawing = picture_xml[drawing_start..drawing_end]
        .replace("r:embed=\"rId1\"", "r:embed=\"rIdCropIntegration\"")
        .replacen("<w:drawing>", "<w:drawing xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">", 1);
    let mut pixels = Vec::new();
    pictures
        .by_name("word/media/bands.png")
        .unwrap()
        .read_to_end(&mut pixels)
        .unwrap();
    let mut package = rewrite_part(SMARTART_REPORT, "word/document.xml", |xml| {
        xml.replacen(
            "</w:p>",
            &format!("</w:p><w:p><w:r>{drawing}</w:r></w:p>"),
            1,
        )
    });
    package = rewrite_part(&package, "word/_rels/document.xml.rels", |xml| {
        xml.replace("</Relationships>", "<Relationship Id=\"rIdCropIntegration\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/crop-integration.png\"/></Relationships>")
    });
    package = rewrite_part(&package, "[Content_Types].xml", |xml| {
        if xml.contains("Extension=\"png\"") {
            xml.into()
        } else {
            xml.replace(
                "</Types>",
                "<Default Extension=\"png\" ContentType=\"image/png\"/></Types>",
            )
        }
    });
    let mut archive = zip::ZipArchive::new(Cursor::new(&package)).unwrap();
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..archive.len() {
        let mut part = archive.by_index(index).unwrap();
        let name = part.name().expect("entry name should decode").into_owned();
        output
            .start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        std::io::copy(&mut part, &mut output).unwrap();
    }
    output
        .start_file(
            "word/media/crop-integration.png",
            zip::write::SimpleFileOptions::default(),
        )
        .unwrap();
    output.write_all(&pixels).unwrap();
    let package = output.finish().unwrap().into_inner();
    let (document, _) = DocxParser
        .parse(&package, &ConvertOptions::default())
        .unwrap();
    assert_eq!(drawing_groups(all_blocks(&document)).len(), 2);
    let pictures: Vec<&crate::ir::ImageData> = all_blocks(&document)
        .iter()
        .filter_map(|block| {
            if let Block::FloatingImage(image) = block {
                Some(&image.image)
            } else {
                None
            }
        })
        .collect();
    assert_eq!(pictures.len(), 1);
    let crop = pictures[0]
        .crop
        .expect("the picture keeps its own source rectangle after the diagram");
    assert_eq!(
        (crop.left, crop.top, crop.right, crop.bottom),
        (0.25, 0.0, 0.25, 0.0)
    );
}

#[test]
fn docx_cached_smartart_reports_a_diagram_in_an_unsupported_hyperlink() {
    let package = rewrite_part(SMARTART_REPORT, "word/document.xml", |xml| {
        let start = xml.find("<w:drawing>").unwrap();
        let end = xml[start..].find("</w:drawing>").unwrap() + start + "</w:drawing>".len();
        let drawing = &xml[start..end];
        xml.replacen("<w:body>", &format!("<w:body><w:p><w:hyperlink w:anchor=\"synthetic\"><w:r>{drawing}</w:r></w:hyperlink></w:p>"), 1)
    });
    let (_, warnings) = DocxParser
        .parse(&package, &ConvertOptions::default())
        .unwrap();
    assert!(
        warnings.iter().any(|warning| matches!(warning,
            ConvertWarning::ParseSkipped { format, reason }
                if format == "DOCX" && reason.contains("SmartArt") && reason.contains("hyperlink")
        )),
        "unsupported visible diagram must not disappear without a diagnostic"
    );
}

#[test]
fn docx_cached_smartart_text_rectangle_is_independent_of_shape_geometry() {
    for (width, height) in [(153.0, 53.55), (100.0, 40.0)] {
        let mut source = zip::ZipArchive::new(Cursor::new(SMARTART_REPORT)).unwrap();
        let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
        for index in 0..source.len() {
            let mut entry = source.by_index(index).unwrap();
            let mut bytes = Vec::new();
            entry.read_to_end(&mut bytes).unwrap();
            if entry.name().expect("entry name should decode") == "word/diagrams/drawing1.xml" {
                let replacement = format!(
                    "cx=\"{}\" cy=\"{}\"",
                    (width * 12700.0) as i64,
                    (height * 12700.0) as i64
                );
                bytes = String::from_utf8(bytes)
                    .unwrap()
                    .replace("cx=\"1943100\" cy=\"680085\"", &replacement)
                    .into_bytes();
            }
            output
                .start_file(
                    entry.name().expect("entry name should decode").into_owned(),
                    zip::write::SimpleFileOptions::default(),
                )
                .unwrap();
            output.write_all(&bytes).unwrap();
        }
        let package = output.finish().unwrap().into_inner();
        let (document, _) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        let drawings = drawing_groups(all_blocks(&document));
        let text = drawings
            .iter()
            .flat_map(|group| &group.elements)
            .find(|element| match &element.kind {
                crate::ir::FixedElementKind::TextBox(text_box) => {
                    text_box.content.iter().any(|block| match block {
                        Block::Paragraph(paragraph) => {
                            paragraph.runs.iter().any(|run| run.text == "Foo")
                        }
                        _ => false,
                    })
                }
                _ => false,
            })
            .expect("cached diagram label Foo");
        assert!((text.width - width).abs() < 0.00001);
        assert!((text.height - height).abs() < 0.00001);
        assert!((text.x - 126.0).abs() < 0.00001);
        assert!(text.y.abs() < 0.00001);
    }
}

#[test]
fn docx_cached_smartart_preserves_run_highlight_colors() {
    for (hex, expected) in [
        ("00FF00", Color::new(0, 255, 0)),
        ("FF0000", Color::new(255, 0, 0)),
        ("114488", Color::new(17, 68, 136)),
    ] {
        let package = rewrite_part(SMARTART_REPORT, "word/diagrams/drawing1.xml", |xml| {
            xml.replace("00FF00", hex)
        });
        let (document, _) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        let groups = drawing_groups(all_blocks(&document));
        let run = groups
            .iter()
            .flat_map(|group| &group.elements)
            .filter_map(|element| match &element.kind {
                crate::ir::FixedElementKind::TextBox(text_box) => Some(&text_box.content),
                _ => None,
            })
            .flatten()
            .filter_map(|block| match block {
                Block::Paragraph(paragraph) => Some(&paragraph.runs),
                _ => None,
            })
            .flatten()
            .find(|run| run.text == "Foo")
            .unwrap();
        assert_eq!(run.style.highlight, Some(expected));
        assert_ne!(run.style.color, Some(expected));
    }
}

#[test]
fn docx_cached_smartart_four_way_arrow_keeps_its_adjustable_outline() {
    let mut outlines = Vec::new();
    for adjustment in [2_000, 7_000] {
        let package = rewrite_part(SMARTART_REPORT, "word/diagrams/drawing2.xml", |xml| {
            xml.replace(
                "name=\"adj1\" fmla=\"val 2000\"",
                &format!("name=\"adj1\" fmla=\"val {adjustment}\""),
            )
        });
        let (document, _) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        let groups = drawing_groups(all_blocks(&document));
        let outline = groups[1]
            .elements
            .iter()
            .find_map(|element| {
                if let crate::ir::FixedElementKind::Shape(shape) = &element.kind
                    && let crate::ir::ShapeKind::Path { subpaths } = &shape.kind
                {
                    return Some(subpaths);
                }
                None
            })
            .expect(
                "four-way arrow must have an outline instead of painting its bounding rectangle",
            );
        assert_eq!(outline.len(), 1);
        assert!(outline[0].closed);
        assert!(outline[0].vertices.len() > 12);
        assert!(!outline[0].vertices.contains(&(0.0, 0.0)));
        outlines.push(outline.clone());
    }
    assert_ne!(
        outlines[0], outlines[1],
        "shaft adjustment must change the painted outline"
    );
}

#[test]
fn docx_cached_smartart_text_rotation_is_relative_to_its_shape() {
    for (shape_rotation, text_rotation, expected) in [
        (10_800_000, 10_800_000, 0.0),
        (20_700_000, -20_700_000, 0.0),
        (1_800_000, 900_000, 45.0),
    ] {
        let package = rewrite_part(SMARTART_REPORT, "word/diagrams/drawing1.xml", |xml| {
            let label = xml.find(">Foo</a:t>").unwrap();
            let start = xml[..label].rfind("<dsp:sp ").unwrap();
            let end = xml[label..].find("</dsp:sp>").unwrap() + label + "</dsp:sp>".len();
            let shape = xml[start..end]
                .replacen("<a:xfrm>", &format!("<a:xfrm rot=\"{shape_rotation}\">"), 1)
                .replacen(
                    "<dsp:txXfrm>",
                    &format!("<dsp:txXfrm rot=\"{text_rotation}\">"),
                    1,
                );
            format!("{}{}{}", &xml[..start], shape, &xml[end..])
        });
        let (document, _) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        let groups = drawing_groups(all_blocks(&document));
        let text = groups[0]
            .elements
            .iter()
            .filter_map(|element| {
                if let crate::ir::FixedElementKind::TextBox(text) = &element.kind {
                    return Some(text);
                }
                None
            })
            .find(|text| {
                text.content.iter().any(|block| matches!(block,
            Block::Paragraph(paragraph) if paragraph.runs.iter().any(|run| run.text == "Foo")
        ))
            })
            .unwrap();
        assert_eq!(
            text.shape_rotation_deg
                .unwrap_or_default()
                .rem_euclid(360.0),
            expected
        );
    }
}

#[test]
fn docx_cached_smartart_degradation_warnings_identify_the_document_format() {
    for preset in ["cube", "unrecognizedPreset"] {
        let package = rewrite_part(SMARTART_REPORT, "word/diagrams/drawing2.xml", |xml| {
            xml.replace("prst=\"quadArrow\"", &format!("prst=\"{preset}\""))
        });
        let (document, warnings) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        assert_eq!(drawing_groups(all_blocks(&document)).len(), 2);
        assert!(
            warnings.iter().any(|warning| matches!(warning,
                ConvertWarning::PartialElement { format, element, .. }
                    if format == "DOCX" && element.contains(preset)
            )),
            "cached diagram warning must retain its DOCX owner: {warnings:?}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_cached_smartart_effect_extents_position_ink_and_reserve_flow() {
    let placement = |left: f64, top: f64, right: f64, bottom: f64| -> (f64, f64, f64, f64) {
        let package = rewrite_part(SMARTART_REPORT, "word/document.xml", |xml| {
            xml.replacen(
                r#"<wp:effectExtent l="12700" t="38100" r="12700" b="12700"/>"#,
                &format!(
                    r#"<wp:effectExtent l="{}" t="{}" r="{}" b="{}"/>"#,
                    (left * 12700.0) as i64,
                    (top * 12700.0) as i64,
                    (right * 12700.0) as i64,
                    (bottom * 12700.0) as i64,
                ),
                1,
            )
            .replacen("</w:p>", "<w:r><w:t>End</w:t></w:r></w:p>", 1)
        });
        let (document, _) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        let source = crate::render::typst_gen::generate_typst(&document)
            .unwrap()
            .source;
        let placed = crate::render::pdf::compiled_text_runs(&source, 0).unwrap();
        let first = placed.iter().find(|run| run.text == "Foo").unwrap();
        let end = placed.iter().find(|run| run.text == "End").unwrap();
        let following = placed
            .iter()
            .filter(|run| run.text == "Dolor")
            .max_by(|a, b| a.baseline_pt.total_cmp(&b.baseline_pt))
            .unwrap();
        (
            first.left_pt,
            first.baseline_pt,
            following.baseline_pt,
            end.left_pt,
        )
    };
    let control = placement(0.0, 0.0, 0.0, 0.0);
    for (left, top, right, bottom) in [
        (1.0, 3.0, 1.0, 1.0),
        (2.5, 7.0, 0.75, 4.0),
        (0.0, 0.0, 0.0, 5.0),
    ] {
        let current = placement(left, top, right, bottom);
        assert!((current.0 - control.0 - left).abs() < 0.01);
        assert!((current.1 - control.1 - top).abs() < 0.01);
        assert!((current.3 - control.3 - left - right).abs() < 0.01);
        assert!(
            (current.2 - control.2 - top - bottom).abs() < 0.01,
            "following diagram must reserve both effect edges: {current:?}, {control:?}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_cached_smartart_auto_spacing_reserves_the_host_run_line_surplus() {
    let package = rewrite_part(SMARTART_REPORT, "word/document.xml", |xml| {
        let first_end = xml.find("</w:p>").unwrap() + "</w:p>".len();
        let next_drawing = first_end + xml[first_end..].find("<w:drawing>").unwrap();
        let next_paragraph = xml[..next_drawing].rfind("<w:p ").unwrap();
        format!("{}{}", &xml[..first_end], &xml[next_paragraph..])
    });
    let measure = |size: f64, proportion: f64, has_text: bool| -> (f64, f64, Option<f64>) {
        let (mut document, _) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        let crate::ir::Page::Flow(page) = &mut document.pages[0] else {
            panic!("SmartArt report must be a flowing document");
        };
        let Block::Paragraph(paragraph) = &mut page.content[0] else {
            panic!("first diagram must remain in its anchor paragraph");
        };
        paragraph.style.line_spacing = Some(LineSpacing::Proportional(proportion));
        paragraph.runs[0].style.font_family = Some("Noto Sans".into());
        paragraph.runs[0].style.font_size = Some(size);
        if has_text {
            paragraph.runs.push(Run {
                text: "X".into(),
                style: paragraph.runs[0].style.clone(),
                href: None,
                footnote: None,
                inline_box: None,
            });
        }
        let source = crate::render::typst_gen::generate_typst(&document)
            .unwrap()
            .source;
        let placed = crate::render::pdf::compiled_text_runs(&source, 0).unwrap();
        let first = placed.iter().find(|run| run.text == "Foo").unwrap();
        let following = placed
            .iter()
            .filter(|run| run.text == "Dolor")
            .max_by(|a, b| a.baseline_pt.total_cmp(&b.baseline_pt))
            .unwrap();
        (
            first.baseline_pt,
            following.baseline_pt,
            placed
                .iter()
                .find(|run| run.text == "X")
                .map(|run| run.baseline_pt),
        )
    };
    let (_, _, natural_pitch_em) = crate::render::pdf::font_line_metrics_em("Noto Sans").unwrap();
    for size in [10.0, 12.0, 24.0] {
        for has_text in [false, true] {
            let control = measure(size, 1.0, has_text);
            for proportion in [278.0 / 240.0, 1.5] {
                let current = measure(size, proportion, has_text);
                let expected_extra = natural_pitch_em * size * (proportion - 1.0);
                assert!((current.0 - control.0).abs() < 0.01);
                assert!(
                    (current.1 - control.1 - expected_extra).abs() < 0.01,
                    "automatic spacing must reserve {expected_extra}pt after the frame: {current:?}, {control:?}"
                );
                if let (Some(current_end), Some(control_end)) = (current.2, control.2) {
                    assert!((current_end - control_end).abs() < 0.01);
                }
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_cached_smartart_auto_spacing_uses_resolved_document_defaults() {
    let consecutive = rewrite_part(SMARTART_REPORT, "word/document.xml", |xml| {
        let first_end = xml.find("</w:p>").unwrap() + "</w:p>".len();
        let next_drawing = first_end + xml[first_end..].find("<w:drawing>").unwrap();
        let next_paragraph = xml[..next_drawing].rfind("<w:p ").unwrap();
        format!("{}{}", &xml[..first_end], &xml[next_paragraph..])
    });
    let measure = |size: u32, line: u32, has_text: bool| -> (f64, f64, Option<f64>) {
        let defaults = rewrite_part(&consecutive, "word/styles.xml", |xml| {
            let start = xml.find("<w:rFonts ").unwrap();
            let end = start + xml[start..].find("/>").unwrap() + 2;
            format!(
                "{}<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/>{}",
                &xml[..start],
                &xml[end..],
            )
            .replacen("w:line=\"278\"", &format!("w:line=\"{line}\""), 1)
            .replacen(
                "<w:sz w:val=\"24\"/>",
                &format!("<w:sz w:val=\"{}\"/>", size * 2),
                1,
            )
        });
        let package = if has_text {
            rewrite_part(&defaults, "word/document.xml", |xml| {
                xml.replacen("</w:p>", "<w:r><w:t>X</w:t></w:r></w:p>", 1)
            })
        } else {
            defaults
        };
        let (document, _) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        let source = crate::render::typst_gen::generate_typst(&document)
            .unwrap()
            .source;
        let placed = crate::render::pdf::compiled_text_runs(&source, 0).unwrap();
        let first = placed.iter().find(|run| run.text == "Foo").unwrap();
        let following = placed
            .iter()
            .filter(|run| run.text == "Dolor")
            .max_by(|a, b| a.baseline_pt.total_cmp(&b.baseline_pt))
            .unwrap();
        (
            first.baseline_pt,
            following.baseline_pt,
            placed
                .iter()
                .find(|run| run.text == "X")
                .map(|run| run.baseline_pt),
        )
    };
    let (_, _, natural_pitch_em) = crate::render::pdf::font_line_metrics_em("Noto Sans").unwrap();
    for size in [10, 12, 24] {
        for has_text in [false, true] {
            let control = measure(size, 240, has_text);
            for line in [278, 360] {
                let current = measure(size, line, has_text);
                let extra = natural_pitch_em * f64::from(size) * (f64::from(line) / 240.0 - 1.0);
                assert!((current.0 - control.0).abs() < 0.01);
                assert!(
                    (current.1 - control.1 - extra).abs() < 0.01,
                    "document defaults must reserve {extra}pt: {current:?}, {control:?}"
                );
                if let (Some(current_end), Some(control_end)) = (current.2, control.2) {
                    assert!((current_end - control_end).abs() < 0.01);
                }
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_cached_smartart_wrapped_frames_use_their_own_run_metrics() {
    let combined = rewrite_part(SMARTART_REPORT, "word/document.xml", |xml| {
        let first_end = xml.find("</w:p>").unwrap() + "</w:p>".len();
        let next_drawing = first_end + xml[first_end..].find("<w:drawing>").unwrap();
        let next_paragraph = xml[..next_drawing].rfind("<w:p ").unwrap();
        let next_content = next_paragraph + xml[next_paragraph..].find('>').unwrap() + 1;
        format!(
            "{}{}",
            &xml[..first_end - "</w:p>".len()],
            &xml[next_content..]
        )
    });
    let measure = |first_size: u32, second_size: u32, second_family: &str| -> f64 {
        let package = rewrite_part(&combined, "word/document.xml", |xml| {
            let mut result = xml.to_string();
            let last = result.rfind("<w:noProof/>").unwrap();
            result.insert_str(last, &format!(
                "<w:rFonts w:ascii=\"{second_family}\" w:hAnsi=\"{second_family}\"/><w:sz w:val=\"{}\"/>", second_size * 2,
            ));
            let first = result.find("<w:noProof/>").unwrap();
            result.insert_str(
                first,
                &format!(
                    "<w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/><w:sz w:val=\"{}\"/>",
                    first_size * 2,
                ),
            );
            result
        });
        let (document, _) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        let source = crate::render::typst_gen::generate_typst(&document)
            .unwrap()
            .source;
        let placed = crate::render::pdf::compiled_text_runs(&source, 0).unwrap();
        placed
            .iter()
            .filter(|run| run.text == "Dolor")
            .max_by(|a, b| a.baseline_pt.total_cmp(&b.baseline_pt))
            .unwrap()
            .baseline_pt
    };
    let control = measure(12, 12, "Noto Sans");
    for family in ["Noto Sans", "Noto Serif"] {
        for size in [10, 24] {
            assert!(
                (measure(12, size, family) - control).abs() < 0.01,
                "a later frame's {size}pt {family} run must not move its origin"
            );
        }
    }
    let (_, _, natural_pitch_em) = crate::render::pdf::font_line_metrics_em("Noto Sans").unwrap();
    let expected = natural_pitch_em * 12.0 * (278.0 / 240.0 - 1.0);
    assert!((measure(24, 12, "Noto Sans") - control - expected).abs() < 0.01);
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_cached_smartart_picture_fill_places_the_source_in_its_destination_rectangle() {
    const PICTURE_REPORT: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/smartart_picture_fill.docx");
    let measure = |fill_rect: &str| {
        let package = rewrite_part(PICTURE_REPORT, "word/diagrams/drawing1.xml", |xml| {
            xml.replacen("<a:fillRect l=\"-25000\" r=\"-25000\"/>", fill_rect, 1)
        });
        let (document, _) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        let generated = crate::render::typst_gen::generate_typst(&document).unwrap();
        crate::render::pdf::compiled_image_bounds(&generated.source, &generated.images).unwrap()
    };
    let control = measure("<a:fillRect/>");
    assert_eq!(control.len(), 5);
    for (xml, left, top, right, bottom) in [
        (
            "<a:fillRect l=\"-25000\" r=\"-25000\"/>",
            -0.25,
            0.0,
            -0.25,
            0.0,
        ),
        (
            "<a:fillRect l=\"-50000\" r=\"-50000\"/>",
            -0.5,
            0.0,
            -0.5,
            0.0,
        ),
        (
            "<a:fillRect l=\"25000\" r=\"25000\"/>",
            0.25,
            0.0,
            0.25,
            0.0,
        ),
        (
            "<a:fillRect l=\"-50000\" t=\"10000\" r=\"25000\" b=\"20000\"/>",
            -0.5,
            0.1,
            0.25,
            0.2,
        ),
        (
            "<a:fillRect l=\"-25000\" r=\"-25000\"></a:fillRect>",
            -0.25,
            0.0,
            -0.25,
            0.0,
        ),
    ] {
        let actual = measure(xml);
        assert_eq!(actual.len(), control.len());
        let c = control[0];
        let expected = (
            c.left_pt + c.width_pt * left,
            c.top_pt + c.height_pt * top,
            c.width_pt * (1.0 - left - right),
            c.height_pt * (1.0 - top - bottom),
        );
        let a = actual[0];
        for (actual, expected) in [
            (a.left_pt, expected.0),
            (a.top_pt, expected.1),
            (a.width_pt, expected.2),
            (a.height_pt, expected.3),
        ] {
            assert!(
                (actual - expected).abs() < 0.01,
                "{xml}: destination expected {expected}, got {actual}"
            );
        }
        assert!(
            a.is_clipped,
            "the expanded picture must remain clipped to its shape frame"
        );
        assert_eq!(
            actual[1..],
            control[1..],
            "unmodified picture fills must not move"
        );
    }
    let empty = measure("<a:fillRect l=\"100000\"/>");
    assert_eq!(empty.len(), 4, "an empty destination must paint no image");
    assert_eq!(empty, control[1..]);
}

#[test]
fn docx_cached_smartart_picture_fill_masks_in_destination_coordinates() {
    const PICTURE_REPORT: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/smartart_picture_fill.docx");
    for (fill_rect, source_rect, transparent, painted) in [
        (
            "<a:fillRect l=\"-25000\" r=\"-25000\"/>",
            "<a:srcRect/>",
            (160, 120),
            (400, 120),
        ),
        (
            "<a:fillRect l=\"25000\" r=\"25000\"/>",
            "<a:srcRect/>",
            (0, 0),
            (80, 60),
        ),
        (
            "<a:fillRect l=\"-25000\" r=\"-25000\"/>",
            "<a:srcRect l=\"20000\" t=\"10000\" r=\"10000\" b=\"20000\"/>",
            (240, 120),
            (400, 120),
        ),
    ] {
        let package = rewrite_part(PICTURE_REPORT, "word/diagrams/drawing1.xml", |xml| {
            xml.replacen("<a:fillRect l=\"-25000\" r=\"-25000\"/>", fill_rect, 1)
                .replacen("<a:srcRect/>", source_rect, 1)
        });
        let (document, warnings) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        assert!(warnings.is_empty(), "unexpected warnings: {warnings:?}");
        let group = drawing_groups(all_blocks(&document))[0];
        let image = group
            .elements
            .iter()
            .find_map(|element| match &element.kind {
                crate::ir::FixedElementKind::Image(image) => Some(image),
                _ => None,
            })
            .unwrap();
        let rgba = image::load_from_memory(&image.data).unwrap().into_rgba8();
        assert_eq!(
            rgba.get_pixel(transparent.0, transparent.1)[3],
            0,
            "outside the destination ellipse: {fill_rect} {source_rect}"
        );
        assert_eq!(
            rgba.get_pixel(painted.0, painted.1)[3],
            255,
            "inside the destination ellipse: {fill_rect} {source_rect}"
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_picture_strokes_follow_the_authored_outline_after_masking() {
    use typst::visualize::{Color as TypstColor, CurveItem, Geometry};
    const PICTURES: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/smartart_picture_fill.docx");
    for (preset, curved) in [("ellipse", true), ("roundRect", true), ("rect", false)] {
        let package = rewrite_part(PICTURES, "word/diagrams/drawing1.xml", |xml| {
            xml.replace("prst=\"ellipse\"", &format!("prst=\"{preset}\""))
        });
        let (document, _) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        let generated = crate::render::typst_gen::generate_typst(&document).unwrap();
        let painted =
            crate::render::pdf::compiled_paint_sequence(&generated.source, &generated.images, 0)
                .unwrap();
        let outlines: Vec<_> = painted
            .iter()
            .filter(|item| {
                item.stroke.as_ref().is_some_and(|stroke| {
                    stroke
                        .color
                        .as_ref()
                        .is_some_and(|color| color.to_rgb() == TypstColor::WHITE.to_rgb())
                        && (stroke.thickness_pt - 2.0).abs() < 0.01
                }) && (item.bounds.2 - item.bounds.0) > 75.0
                    && (item.bounds.2 - item.bounds.0) < 85.0
                    && (item.bounds.3 - item.bounds.1) > 75.0
                    && (item.bounds.3 - item.bounds.1) < 90.0
            })
            .collect();
        assert_eq!(
            outlines.len(),
            5,
            "{preset}: five authored picture outlines"
        );
        for outline in outlines {
            if let Some(Geometry::Curve(curve)) = &outline.geometry {
                let lines = curve
                    .0
                    .iter()
                    .filter(|item| matches!(item, CurveItem::Line(_)))
                    .count();
                assert_eq!(
                    lines > 0,
                    preset == "roundRect",
                    "{preset}: straight sides distinguish rounded rectangles from ellipses"
                );
            }
            assert_eq!(
                matches!(outline.geometry, Some(Geometry::Curve(_))),
                curved,
                "{preset}: compiled outline must preserve the authored geometry: {outline:?}"
            );
        }
        let pdf = crate::convert_bytes(&package, crate::Format::Docx, &ConvertOptions::default())
            .unwrap();
        let text = pdf_extract::extract_text_from_mem(&pdf.pdf).unwrap();
        for label in ["Foo", "Bar", "Lorem", "Ipsum", "Dolor"] {
            assert!(
                text.contains(label),
                "{preset}: missing searchable label {label}"
            );
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_picture_strokes_keep_custom_paths_alpha_and_frame_transforms() {
    use crate::render::pdf::{PaintedKind, compiled_paint_sequence};
    use typst::visualize::{CurveItem, Geometry};
    const PICTURES: &[u8] =
        include_bytes!("../../../../tests/fixtures/docx/docxide/smartart_picture_fill.docx");
    const TRIANGLE: &str = "<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l=\"0\" t=\"0\" r=\"100000\" b=\"100000\"/><a:pathLst><a:path w=\"100000\" h=\"100000\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo><a:lnTo><a:pt x=\"100000\" y=\"0\"/></a:lnTo><a:lnTo><a:pt x=\"0\" y=\"100000\"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>";
    for (transform, alpha, custom) in [
        ("rot=\"1800000\" flipH=\"1\"", 25000, false),
        ("rot=\"-4200000\" flipV=\"1\"", 65000, true),
        ("flipH=\"1\" flipV=\"1\"", 100000, true),
    ] {
        let package = rewrite_part(PICTURES, "word/diagrams/drawing1.xml", |xml| {
            let mut xml = xml
                .replace("<a:xfrm>", &format!("<a:xfrm {transform}>"))
                .replace("<a:fillRect l=\"-25000\" r=\"-25000\"/>", "<a:fillRect/>")
                .replace(
                    "<a:schemeClr val=\"lt1\">",
                    &format!("<a:schemeClr val=\"lt1\"><a:alpha val=\"{alpha}\"/>"),
                );
            if custom {
                xml = xml.replace(
                    "<a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom>",
                    TRIANGLE,
                );
            }
            xml
        });
        let (document, _) = DocxParser
            .parse(&package, &ConvertOptions::default())
            .unwrap();
        let generated = crate::render::typst_gen::generate_typst(&document).unwrap();
        let painted = compiled_paint_sequence(&generated.source, &generated.images, 0).unwrap();
        let images: Vec<_> = painted
            .iter()
            .filter(|item| item.kind == PaintedKind::Image)
            .collect();
        assert_eq!(images.len(), 5);
        let outlines: Vec<_> = painted
            .iter()
            .filter(|item| {
                let Some(Geometry::Curve(curve)) = &item.geometry else {
                    return false;
                };
                let size = curve.bbox(None);
                item.stroke
                    .as_ref()
                    .is_some_and(|stroke| (stroke.thickness_pt - 2.0).abs() < 0.01)
                    && size.max.x.to_pt() - size.min.x.to_pt() > 75.0
                    && size.max.x.to_pt() - size.min.x.to_pt() < 85.0
                    && size.max.y.to_pt() - size.min.y.to_pt() > 75.0
                    && size.max.y.to_pt() - size.min.y.to_pt() < 90.0
            })
            .collect();
        assert_eq!(
            outlines.len(),
            5,
            "{transform}: every picture has a curved/path outline"
        );
        for (image, outline) in images.iter().zip(outlines) {
            let color = outline
                .stroke
                .as_ref()
                .unwrap()
                .color
                .as_ref()
                .unwrap()
                .to_rgb();
            assert!(
                (color.alpha as f64 - alpha as f64 / 100000.0).abs() < 0.005,
                "{transform}: source line alpha must survive masking"
            );
            let a = image.transform;
            let b = outline.transform;
            let Some(Geometry::Curve(curve)) = &outline.geometry else {
                unreachable!();
            };
            let bounds = curve.bbox(None);
            if custom {
                let mut vertices: Vec<_> = curve
                    .0
                    .iter()
                    .filter_map(|item| match item {
                        CurveItem::Move(point) | CurveItem::Line(point) => Some(*point),
                        CurveItem::Cubic(..) | CurveItem::Close => None,
                    })
                    .collect();
                if vertices.first() == vertices.last() {
                    vertices.pop();
                }
                assert_eq!(
                    vertices.len(),
                    3,
                    "custom triangle must preserve three distinct vertices"
                );
                assert!(
                    curve.0.iter().any(|item| matches!(item, CurveItem::Close)),
                    "custom picture stroke must close its triangle"
                );
                for (actual, expected) in vertices.iter().zip([
                    bounds.min,
                    typst::layout::Point::new(bounds.max.x, bounds.min.y),
                    typst::layout::Point::new(bounds.min.x, bounds.max.y),
                ]) {
                    assert!(
                        (actual.x.to_pt() - expected.x.to_pt()).abs() < 0.01
                            && (actual.y.to_pt() - expected.y.to_pt()).abs() < 0.01,
                        "custom picture stroke must retain the asymmetric triangle outline"
                    );
                }
            }
            // Ellipses have a centre-based internal curve; compare painted
            // positions rather than requiring identical internal origins.
            let x = (bounds.min.x.to_pt() + bounds.max.x.to_pt()) / 2.0;
            let y = (bounds.min.y.to_pt() + bounds.max.y.to_pt()) / 2.0;
            let stroke_centre = (
                b.sx.get() * x + b.kx.get() * y + b.tx.to_pt(),
                b.ky.get() * x + b.sy.get() * y + b.ty.to_pt(),
            );
            let image_centre = (
                (image.bounds.0 + image.bounds.2) / 2.0,
                (image.bounds.1 + image.bounds.3) / 2.0,
            );
            assert!(
                (stroke_centre.0 - image_centre.0).abs() < 0.01
                    && (stroke_centre.1 - image_centre.1).abs() < 0.01,
                "{transform}: stroke and picture frame centres must coincide"
            );
            for (a, b) in [
                (a.sx.get(), b.sx.get()),
                (a.sy.get(), b.sy.get()),
                (a.kx.get(), b.kx.get()),
                (a.ky.get(), b.ky.get()),
            ] {
                assert!(
                    (a - b).abs() < 0.01,
                    "{transform}: picture and stroke must share page-space transforms, {a} vs {b}"
                );
            }
        }
    }
}
