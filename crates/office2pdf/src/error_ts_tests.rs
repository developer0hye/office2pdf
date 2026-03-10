use super::*;
use ts_rs::TS;

fn cfg() -> ts_rs::Config {
    ts_rs::Config::new()
}

#[test]
fn test_convert_warning_ts_declaration() {
    let decl = ConvertWarning::decl(&cfg());
    assert!(
        decl.contains("ConvertWarning"),
        "ConvertWarning TS decl: {decl}"
    );
    assert!(
        decl.contains("UnsupportedElement"),
        "should contain UnsupportedElement variant: {decl}"
    );
    assert!(
        decl.contains("PartialElement"),
        "should contain PartialElement variant: {decl}"
    );
}

#[test]
fn test_convert_metrics_ts_declaration() {
    let decl = ConvertMetrics::decl(&cfg());
    assert!(
        decl.contains("ConvertMetrics"),
        "ConvertMetrics TS decl: {decl}"
    );
    assert!(
        decl.contains("page_count"),
        "should contain page_count field: {decl}"
    );
    assert!(
        decl.contains("number"),
        "numeric fields should be number type: {decl}"
    );
}

#[test]
fn test_convert_warning_ts_export() {
    let ts = ConvertWarning::export_to_string(&cfg()).unwrap();
    assert!(ts.contains("ConvertWarning"));
}

#[test]
fn test_convert_metrics_ts_export() {
    let ts = ConvertMetrics::export_to_string(&cfg()).unwrap();
    assert!(ts.contains("ConvertMetrics"));
}
