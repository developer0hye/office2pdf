use super::*;
use ts_rs::TS;

fn cfg() -> ts_rs::Config {
    ts_rs::Config::new()
}

#[test]
fn test_format_ts_declaration() {
    let decl = Format::decl(&cfg());
    assert!(decl.contains("Format"), "Format TS decl: {decl}");
    assert!(decl.contains("Docx"), "should contain Docx variant");
    assert!(decl.contains("Pptx"), "should contain Pptx variant");
    assert!(decl.contains("Xlsx"), "should contain Xlsx variant");
}

#[test]
fn test_paper_size_ts_declaration() {
    let decl = PaperSize::decl(&cfg());
    assert!(decl.contains("PaperSize"), "PaperSize TS decl: {decl}");
    assert!(decl.contains("A4"), "should contain A4 variant");
    assert!(decl.contains("Letter"), "should contain Letter variant");
    assert!(decl.contains("Legal"), "should contain Legal variant");
    assert!(decl.contains("Custom"), "should contain Custom variant");
}

#[test]
fn test_pdf_standard_ts_declaration() {
    let decl = PdfStandard::decl(&cfg());
    assert!(decl.contains("PdfStandard"), "PdfStandard TS decl: {decl}");
    assert!(decl.contains("PdfA2b"), "should contain PdfA2b variant");
}

#[test]
fn test_slide_range_ts_declaration() {
    let decl = SlideRange::decl(&cfg());
    assert!(decl.contains("SlideRange"), "SlideRange TS decl: {decl}");
    assert!(decl.contains("start"), "should contain start field");
    assert!(decl.contains("end"), "should contain end field");
    assert!(decl.contains("number"), "fields should be number type");
}

#[test]
fn test_convert_options_ts_declaration() {
    let decl = ConvertOptions::decl(&cfg());
    assert!(
        decl.contains("ConvertOptions"),
        "ConvertOptions TS decl: {decl}"
    );
    assert!(
        decl.contains("tagged"),
        "should contain tagged field: {decl}"
    );
    assert!(
        decl.contains("pdf_ua"),
        "should contain pdf_ua field: {decl}"
    );
}

#[test]
fn test_format_ts_export() {
    let ts = Format::export_to_string(&cfg()).unwrap();
    assert!(ts.contains("Format"));
}

#[test]
fn test_convert_options_ts_export() {
    let ts = ConvertOptions::export_to_string(&cfg()).unwrap();
    assert!(ts.contains("ConvertOptions"));
}
