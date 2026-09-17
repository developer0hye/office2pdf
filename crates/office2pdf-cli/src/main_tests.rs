use super::*;
use std::io::Cursor;

fn make_test_docx() -> Vec<u8> {
    let docx = docx_rs::Docx::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello batch")),
    );
    let mut buf = Cursor::new(Vec::new());
    docx.build().pack(&mut buf).unwrap();
    buf.into_inner()
}

// --- Unit tests for determine_output_path ---

#[test]
fn test_determine_output_path_default() {
    let input = PathBuf::from("/tmp/report.docx");
    let result = determine_output_path(&input, None, None);
    assert_eq!(result, PathBuf::from("/tmp/report.pdf"));
}

#[test]
fn test_determine_output_path_with_output() {
    let input = PathBuf::from("/tmp/report.docx");
    let output = PathBuf::from("/custom/output.pdf");
    let result = determine_output_path(&input, Some(&output), None);
    assert_eq!(result, PathBuf::from("/custom/output.pdf"));
}

#[test]
fn test_determine_output_path_with_outdir() {
    let input = PathBuf::from("/tmp/report.docx");
    let outdir = PathBuf::from("/output");
    let result = determine_output_path(&input, None, Some(&outdir));
    assert_eq!(result, PathBuf::from("/output/report.pdf"));
}

#[test]
fn test_determine_output_path_outdir_replaces_extension() {
    let input = PathBuf::from("/data/slides.pptx");
    let outdir = PathBuf::from("/pdfs");
    let result = determine_output_path(&input, None, Some(&outdir));
    assert_eq!(result, PathBuf::from("/pdfs/slides.pdf"));
}

// --- Integration tests for batch conversion ---

#[test]
fn test_batch_convert_multiple_files() {
    let dir = std::env::temp_dir().join("office2pdf_batch_test_multi");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();

    let docx_data = make_test_docx();
    let file1 = dir.join("doc1.docx");
    let file2 = dir.join("doc2.docx");
    std::fs::write(&file1, &docx_data).unwrap();
    std::fs::write(&file2, &docx_data).unwrap();

    let inputs = vec![file1, file2];
    let options = ConvertOptions::default();
    let result = convert_batch(&inputs, None, &options, false, 1);

    assert_eq!(result.succeeded.len(), 2);
    assert_eq!(result.failed.len(), 0);
    assert!(dir.join("doc1.pdf").exists());
    assert!(dir.join("doc2.pdf").exists());

    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_batch_convert_partial_failure() {
    let dir = std::env::temp_dir().join("office2pdf_batch_test_fail");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();

    let docx_data = make_test_docx();
    let file1 = dir.join("good.docx");
    let file2 = dir.join("bad.txt");
    std::fs::write(&file1, &docx_data).unwrap();
    std::fs::write(&file2, b"not a valid document").unwrap();

    let inputs = vec![file1, file2.clone()];
    let options = ConvertOptions::default();
    let result = convert_batch(&inputs, None, &options, false, 1);

    assert_eq!(result.succeeded.len(), 1);
    assert_eq!(result.failed.len(), 1);
    assert!(dir.join("good.pdf").exists());
    assert_eq!(result.failed[0].0, file2);

    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_batch_convert_with_outdir() {
    let dir = std::env::temp_dir().join("office2pdf_batch_test_outdir");
    let outdir = dir.join("output");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    std::fs::create_dir_all(&outdir).unwrap();

    let docx_data = make_test_docx();
    let file1 = dir.join("report.docx");
    let file2 = dir.join("memo.docx");
    std::fs::write(&file1, &docx_data).unwrap();
    std::fs::write(&file2, &docx_data).unwrap();

    let inputs = vec![file1, file2];
    let options = ConvertOptions::default();
    let result = convert_batch(&inputs, Some(&outdir), &options, false, 1);

    assert_eq!(result.succeeded.len(), 2);
    assert_eq!(result.failed.len(), 0);
    assert!(outdir.join("report.pdf").exists());
    assert!(outdir.join("memo.pdf").exists());
    // Original directory should NOT have PDFs
    assert!(!dir.join("report.pdf").exists());
    assert!(!dir.join("memo.pdf").exists());

    let _ = std::fs::remove_dir_all(&dir);
}

// --- Parallel batch conversion tests ---

#[test]
fn test_batch_convert_parallel_jobs_2() {
    let dir = std::env::temp_dir().join("office2pdf_parallel_test_j2");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();

    let docx_data = make_test_docx();
    let inputs: Vec<PathBuf> = (0..4)
        .map(|i| {
            let path = dir.join(format!("doc{i}.docx"));
            std::fs::write(&path, &docx_data).unwrap();
            path
        })
        .collect();

    let options = ConvertOptions::default();
    let result = convert_batch(&inputs, None, &options, false, 2);

    assert_eq!(result.succeeded.len(), 4);
    assert_eq!(result.failed.len(), 0);
    for i in 0..4 {
        let pdf_path = dir.join(format!("doc{i}.pdf"));
        assert!(pdf_path.exists(), "doc{i}.pdf should exist");
        let pdf_bytes = std::fs::read(&pdf_path).unwrap();
        assert!(pdf_bytes.len() > 100, "PDF should have real content");
    }

    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_batch_convert_parallel_partial_failure() {
    let dir = std::env::temp_dir().join("office2pdf_parallel_fail_test");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();

    let docx_data = make_test_docx();
    let good = dir.join("good.docx");
    let bad = dir.join("bad.txt");
    std::fs::write(&good, &docx_data).unwrap();
    std::fs::write(&bad, b"not a valid document").unwrap();

    let inputs = vec![good, bad.clone()];
    let options = ConvertOptions::default();
    let result = convert_batch(&inputs, None, &options, false, 2);

    assert_eq!(result.succeeded.len(), 1);
    assert_eq!(result.failed.len(), 1);
    assert!(dir.join("good.pdf").exists());
    assert_eq!(result.failed[0].0, bad);

    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_batch_convert_parallel_with_outdir() {
    let dir = std::env::temp_dir().join("office2pdf_parallel_outdir_test");
    let outdir = dir.join("output");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    std::fs::create_dir_all(&outdir).unwrap();

    let docx_data = make_test_docx();
    let inputs: Vec<PathBuf> = (0..3)
        .map(|i| {
            let path = dir.join(format!("file{i}.docx"));
            std::fs::write(&path, &docx_data).unwrap();
            path
        })
        .collect();

    let options = ConvertOptions::default();
    let result = convert_batch(&inputs, Some(&outdir), &options, false, 2);

    assert_eq!(result.succeeded.len(), 3);
    assert_eq!(result.failed.len(), 0);
    for i in 0..3 {
        assert!(outdir.join(format!("file{i}.pdf")).exists());
        // Original directory should NOT have PDFs
        assert!(!dir.join(format!("file{i}.pdf")).exists());
    }

    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_batch_convert_single_file_with_jobs() {
    // Single file should work fine even with jobs > 1
    let dir = std::env::temp_dir().join("office2pdf_parallel_single_test");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();

    let docx_data = make_test_docx();
    let input = dir.join("single.docx");
    std::fs::write(&input, &docx_data).unwrap();

    let inputs = vec![input];
    let options = ConvertOptions::default();
    let result = convert_batch(&inputs, None, &options, false, 4);

    assert_eq!(result.succeeded.len(), 1);
    assert_eq!(result.failed.len(), 0);
    assert!(dir.join("single.pdf").exists());

    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_batch_convert_sequential_jobs_1() {
    // jobs=1 should use sequential path
    let dir = std::env::temp_dir().join("office2pdf_sequential_test");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();

    let docx_data = make_test_docx();
    let inputs: Vec<PathBuf> = (0..3)
        .map(|i| {
            let path = dir.join(format!("seq{i}.docx"));
            std::fs::write(&path, &docx_data).unwrap();
            path
        })
        .collect();

    let options = ConvertOptions::default();
    let result = convert_batch(&inputs, None, &options, false, 1);

    assert_eq!(result.succeeded.len(), 3);
    assert_eq!(result.failed.len(), 0);

    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_convert_single_with_metrics() {
    let dir = std::env::temp_dir().join("office2pdf_metrics_test");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();

    let docx_data = make_test_docx();
    let input = dir.join("report.docx");
    let output = dir.join("report.pdf");
    std::fs::write(&input, &docx_data).unwrap();

    let options = ConvertOptions::default();
    // Should succeed with metrics=true (metrics printed to stderr)
    convert_single(&input, &output, &options, true).unwrap();
    assert!(output.exists());

    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn fixed_raster_image_does_not_round_below_the_exact_bottom_edge() {
    use lopdf::{Object, content::Content};
    use office2pdf::ir::{
        Document as OfficeDocument, FixedElement, FixedElementKind, FixedPage, ImageData,
        ImageFormat, Metadata, Page, PageSize, StyleSheet,
    };

    // A 324pt top plus 183.6pt height on a 540pt slide leaves an exact
    // 32.4pt PDF-space bottom. The f32 PDF transform used by Typst rounds the
    // unadjusted subtraction to 32.399994pt, which makes a 150-DPI renderer
    // blend the bottom source row into a second device row (issue #666).
    const RED_PIXEL_PNG: &[u8] = &[
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44,
        0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90,
        0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8,
        0xcf, 0xc0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xdd, 0x8d, 0xb0, 0x00, 0x00, 0x00,
        0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
    ];
    let fixed_image = |y| FixedElement {
        x: 216.0,
        y,
        width: 403.2,
        height: 183.6,
        kind: FixedElementKind::Image(ImageData {
            data: RED_PIXEL_PNG.to_vec(),
            rotation_deg: None,
            flip_h: false,
            flip_v: false,
            format: ImageFormat::Png,
            width: Some(403.2),
            height: Some(183.6),
            crop: None,
            stroke: None,
            alignment: None,
            clip_shape: None,
            shadow: None,
            paragraph_spacing: None,
        }),
    };
    let document = OfficeDocument {
        metadata: Metadata::default(),
        pages: vec![Page::Fixed(FixedPage {
            size: PageSize {
                width: 960.0,
                height: 540.0,
            },
            // The first coordinate rounds down at the compiler boundary. The
            // second already rounds up and must stay unchanged, triangulating
            // the conditional tie-break rather than a blanket image shift.
            elements: vec![fixed_image(324.0), fixed_image(0.2)],
            background_color: None,
            background_gradient: None,
        })],
        styles: StyleSheet::default(),
    };

    let pdf = office2pdf::render_document(&document).expect("fixed raster should render");
    let parsed = lopdf::Document::load_mem(&pdf).expect("rendered PDF should parse");
    let page_id = parsed.get_pages()[&1];
    let content =
        Content::decode(&parsed.get_page_content(page_id)).expect("page content should decode");
    let matrices = content
        .operations
        .windows(2)
        .filter_map(|operations| {
            (operations[0].operator == "cm" && operations[1].operator == "Do")
                .then_some(&operations[0].operands)
        })
        .collect::<Vec<_>>();
    assert_eq!(matrices.len(), 2, "each image should have one draw matrix");
    let number = |object: &Object| match object {
        Object::Integer(value) => *value as f64,
        Object::Real(value) => f64::from(*value),
        other => panic!("expected PDF number, got {other:?}"),
    };
    let emitted_bottom = number(&matrices[0][5]);
    let exact_bottom = 540.0 - 324.0 - 183.6;

    assert!(
        emitted_bottom >= exact_bottom,
        "fixed raster bottom rounded below the exact edge: {emitted_bottom} < {exact_bottom}"
    );
    let already_safe_bottom = (540.0_f32 - (0.2_f32 + 183.6_f32)) as f64;
    assert_eq!(
        number(&matrices[1][5]),
        already_safe_bottom,
        "a raster whose transform already rounds upward must not move"
    );
}

#[test]
fn rotated_preset_keeps_explicit_vertical_body_direction_and_top_anchor() {
    use lopdf::{Object, content::Content};

    let fixture = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/pptx/libreoffice/shape-text-rotate.pptx");
    let rendered = office2pdf::convert(&fixture).expect("fixture should render");
    let parsed = lopdf::Document::load_mem(&rendered.pdf).expect("rendered PDF should parse");
    let page_id = parsed.get_pages()[&1];
    let content =
        Content::decode(&parsed.get_page_content(page_id)).expect("page content should decode");
    let number = |object: &Object| match object {
        Object::Integer(value) => *value as f64,
        Object::Real(value) => f64::from(*value),
        other => panic!("expected PDF number, got {other:?}"),
    };
    let text_matrix = content
        .operations
        .windows(8)
        .find_map(|operations| {
            (operations[0].operator == "cm"
                && operations[1].operator == "cs"
                && operations[2].operator == "scn"
                && operations[3].operator == "BT"
                && operations[5].operator == "Tf")
                .then_some(&operations[0].operands)
        })
        .expect("vertical text should have a graphics transform");

    assert!(
        number(&text_matrix[1]) > 0.99 && number(&text_matrix[2]) > 0.99,
        "vert text must keep PowerPoint's bottom-to-top matrix even though its preset shape is rotated: {text_matrix:?}"
    );
    // The final glyph origin includes the platform's resolved font metrics;
    // its x coordinate varies by up to 1.65pt, while Linux differs from macOS
    // by about 0.33pt on the y axis that carries this transformed top anchor.
    // The parser's rotated-pentagon regression checks both overlay axes
    // exactly; this 1pt PDF-level y band remains far below the 24.5pt
    // full-box displacement.
    assert!(
        (number(&text_matrix[5]) - 313.600_28).abs() < 1.0,
        "anchor=t must position vertical text in the pentagon's transformed text rectangle: {text_matrix:?}"
    );
}

// --- PDF merge/split CLI tests ---

fn make_test_pdf(num_pages: u32) -> Vec<u8> {
    use lopdf::{Document, Object, Stream, dictionary};

    let mut doc = Document::with_version("1.7");
    let pages_id = doc.new_object_id();
    let mut page_ids = Vec::new();

    for i in 0..num_pages {
        let content = format!("BT /F1 12 Tf 100 700 Td (Page {}) Tj ET", i + 1);
        let content_id = doc.add_object(Stream::new(dictionary! {}, content.into_bytes()));
        let page_id = doc.add_object(dictionary! {
            "Type" => "Page",
            "Parent" => pages_id,
            "MediaBox" => vec![0.into(), 0.into(), 595.into(), 842.into()],
            "Contents" => content_id,
        });
        page_ids.push(page_id);
    }

    let page_refs: Vec<Object> = page_ids.iter().map(|id| Object::Reference(*id)).collect();

    doc.objects.insert(
        pages_id,
        Object::Dictionary(dictionary! {
            "Type" => "Pages",
            "Count" => num_pages as i64,
            "Kids" => page_refs,
        }),
    );

    let catalog_id = doc.add_object(dictionary! {
        "Type" => "Catalog",
        "Pages" => pages_id,
    });
    doc.trailer.set("Root", Object::Reference(catalog_id));

    let mut output = Vec::new();
    doc.save_to(&mut output).unwrap();
    output
}

#[test]
fn test_cli_merge_command() {
    let dir = std::env::temp_dir().join("office2pdf_cli_merge_test");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();

    let pdf1 = make_test_pdf(1);
    let pdf2 = make_test_pdf(2);
    let file1 = dir.join("a.pdf");
    let file2 = dir.join("b.pdf");
    let output = dir.join("merged.pdf");
    std::fs::write(&file1, &pdf1).unwrap();
    std::fs::write(&file2, &pdf2).unwrap();

    let cmd = Commands::Merge {
        files: vec![file1, file2],
        output: output.clone(),
    };
    handle_command(cmd).unwrap();

    assert!(output.exists());
    let merged_data = std::fs::read(&output).unwrap();
    assert_eq!(pdf_ops::page_count(&merged_data).unwrap(), 3);

    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_cli_split_command() {
    let dir = std::env::temp_dir().join("office2pdf_cli_split_test");
    let outdir = dir.join("splits");
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();

    let pdf = make_test_pdf(4);
    let input = dir.join("doc.pdf");
    std::fs::write(&input, &pdf).unwrap();

    let cmd = Commands::Split {
        input: input.clone(),
        pages: vec!["1-2".to_string(), "3-4".to_string()],
        outdir: outdir.clone(),
    };
    handle_command(cmd).unwrap();

    assert!(outdir.join("doc_pages_1-2.pdf").exists());
    assert!(outdir.join("doc_pages_3-4.pdf").exists());

    let part1 = std::fs::read(outdir.join("doc_pages_1-2.pdf")).unwrap();
    let part2 = std::fs::read(outdir.join("doc_pages_3-4.pdf")).unwrap();
    assert_eq!(pdf_ops::page_count(&part1).unwrap(), 2);
    assert_eq!(pdf_ops::page_count(&part2).unwrap(), 2);

    let _ = std::fs::remove_dir_all(&dir);
}

#[test]
fn test_cli_include_hidden_slides_defaults_to_false() {
    let cli = Cli::try_parse_from(["office2pdf", "slides.pptx"]).unwrap();
    assert!(!cli.include_hidden_slides);
}

#[test]
fn test_cli_include_hidden_slides_flag_sets_option() {
    let cli =
        Cli::try_parse_from(["office2pdf", "--include-hidden-slides", "slides.pptx"]).unwrap();
    let options = build_convert_options(&cli, None, None, None);
    assert!(options.include_hidden_slides);
}
