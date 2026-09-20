#![cfg(not(target_arch = "wasm32"))]
//! Native conversion coverage for both settings of `embedded-fonts`.

use std::path::PathBuf;

use office2pdf::config::ConvertOptions;

fn assert_conversion(relative_path: &str, expected_text: &str) {
    let source = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/golden_mocks/business/sources")
        .join(relative_path);
    // A caller can supply an OFL face even on hosts with no installed fonts.
    let options = ConvertOptions {
        font_bytes: vec![include_bytes!("../fonts/NotoSans-Regular.ttf").to_vec()],
        last_resort_font_family: Some("Noto Sans".into()),
        ..Default::default()
    };
    let result = office2pdf::convert_with_options(source, &options)
        .expect("conversion must succeed with or without Typst's bundled fonts");
    let text = pdf_extract::extract_text_from_mem(&result.pdf).expect("PDF must be readable");
    assert!(
        text.contains(expected_text),
        "expected {expected_text:?} in {text:?}"
    );
}

#[test]
fn docx_conversion_with_caller_font() {
    assert_conversion("docx/01_invoice_en.docx", "Invoice");
}

#[test]
fn xlsx_conversion_with_caller_font() {
    assert_conversion("xlsx/03_inventory_en.xlsx", "Inventory");
}

#[test]
fn pptx_conversion_with_caller_font() {
    assert_conversion("pptx/01_startup_pitch_en.pptx", "Renderly");
}
