use ts_rs::TS;

use crate::config::{ConvertOptions, Format, PaperSize, PdfStandard, SlideRange};
use crate::error::{ConvertMetrics, ConvertWarning};

fn cfg_for_bindings() -> ts_rs::Config {
    let bindings_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("bindings");
    std::fs::create_dir_all(&bindings_dir).unwrap();
    ts_rs::Config::new().with_out_dir(bindings_dir)
}

#[test]
fn test_export_all_types_to_bindings() {
    let cfg = cfg_for_bindings();
    let bindings_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("bindings");

    Format::export_all(&cfg).unwrap();
    PaperSize::export_all(&cfg).unwrap();
    PdfStandard::export_all(&cfg).unwrap();
    SlideRange::export_all(&cfg).unwrap();
    ConvertOptions::export_all(&cfg).unwrap();
    ConvertWarning::export_all(&cfg).unwrap();
    ConvertMetrics::export_all(&cfg).unwrap();

    assert!(bindings_dir.join("Format.ts").exists());
    assert!(bindings_dir.join("PaperSize.ts").exists());
    assert!(bindings_dir.join("PdfStandard.ts").exists());
    assert!(bindings_dir.join("SlideRange.ts").exists());
    assert!(bindings_dir.join("ConvertOptions.ts").exists());
    assert!(bindings_dir.join("ConvertWarning.ts").exists());
    assert!(bindings_dir.join("ConvertMetrics.ts").exists());
}

#[test]
fn test_generated_types_contain_expected_content() {
    let cfg = cfg_for_bindings();
    let bindings_dir = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("bindings");

    Format::export_all(&cfg).unwrap();
    ConvertOptions::export_all(&cfg).unwrap();

    let format_ts = std::fs::read_to_string(bindings_dir.join("Format.ts")).unwrap();
    assert!(
        format_ts.contains("Docx"),
        "Format.ts should contain Docx: {format_ts}"
    );
    assert!(
        format_ts.contains("Pptx"),
        "Format.ts should contain Pptx: {format_ts}"
    );
    assert!(
        format_ts.contains("Xlsx"),
        "Format.ts should contain Xlsx: {format_ts}"
    );

    let opts_ts = std::fs::read_to_string(bindings_dir.join("ConvertOptions.ts")).unwrap();
    assert!(
        opts_ts.contains("tagged"),
        "ConvertOptions.ts should contain tagged: {opts_ts}"
    );
    assert!(
        opts_ts.contains("pdf_ua"),
        "ConvertOptions.ts should contain pdf_ua: {opts_ts}"
    );
    assert!(
        opts_ts.contains("boolean"),
        "boolean fields should be mapped: {opts_ts}"
    );
}
