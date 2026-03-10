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
