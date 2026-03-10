use super::*;

#[test]
fn test_empty_metrics_render() {
    let store = MetricsStore::new();
    let output = store.render();
    assert!(output.contains("# HELP office2pdf_conversions_total"));
    assert!(output.contains("# TYPE office2pdf_conversions_total counter"));
    assert!(output.contains("# HELP office2pdf_active_conversions"));
    assert!(output.contains("# TYPE office2pdf_active_conversions gauge"));
    assert!(output.contains("office2pdf_active_conversions 0"));
}

#[test]
fn test_record_success_increments_counter() {
    let store = MetricsStore::new();
    store.record_success("docx", 0.5, 1024, 2048, 3);
    store.record_success("docx", 0.3, 512, 1024, 1);

    let output = store.render();
    assert!(output.contains("office2pdf_conversions_total{format=\"docx\",status=\"success\"} 2"));
}

#[test]
fn test_record_failure_increments_counters() {
    let store = MetricsStore::new();
    store.record_failure("pptx", "conversion");

    let output = store.render();
    assert!(output.contains("office2pdf_conversions_total{format=\"pptx\",status=\"failure\"} 1"));
    assert!(
        output.contains("office2pdf_errors_total{format=\"pptx\",error_type=\"conversion\"} 1")
    );
}

#[test]
fn test_active_gauge_increment_decrement() {
    let store = MetricsStore::new();
    assert_eq!(store.active.load(Ordering::Relaxed), 0);

    store.start_conversion();
    assert_eq!(store.active.load(Ordering::Relaxed), 1);

    store.start_conversion();
    assert_eq!(store.active.load(Ordering::Relaxed), 2);

    store.end_conversion();
    assert_eq!(store.active.load(Ordering::Relaxed), 1);

    store.end_conversion();
    assert_eq!(store.active.load(Ordering::Relaxed), 0);
}

#[test]
fn test_active_gauge_renders_correctly() {
    let store = MetricsStore::new();
    store.start_conversion();
    let output = store.render();
    assert!(output.contains("office2pdf_active_conversions 1"));
}

#[test]
fn test_duration_histogram_buckets() {
    let store = MetricsStore::new();
    // 50ms = 0.05s, should fall in le=0.05 bucket and above
    store.record_success("docx", 0.05, 100, 200, 1);

    let output = store.render();
    // Should be in le=0.05 bucket
    assert!(
        output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"0.05\"} 1"
        )
    );
    // Should NOT be in le=0.01 bucket
    assert!(
        output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"0.01\"} 0"
        )
    );
    // Should be in +Inf bucket
    assert!(
        output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"+Inf\"} 1"
        )
    );
    // Sum and count
    assert!(output.contains("office2pdf_conversion_duration_seconds_sum{format=\"docx\"} 0.05"));
    assert!(output.contains("office2pdf_conversion_duration_seconds_count{format=\"docx\"} 1"));
}

#[test]
fn test_multiple_formats_tracked_separately() {
    let store = MetricsStore::new();
    store.record_success("docx", 0.1, 100, 200, 1);
    store.record_success("xlsx", 0.2, 300, 400, 2);
    store.record_failure("pptx", "conversion");

    let output = store.render();
    assert!(output.contains("office2pdf_conversions_total{format=\"docx\",status=\"success\"} 1"));
    assert!(output.contains("office2pdf_conversions_total{format=\"xlsx\",status=\"success\"} 1"));
    assert!(output.contains("office2pdf_conversions_total{format=\"pptx\",status=\"failure\"} 1"));
}

#[test]
fn test_histogram_cumulative_counts() {
    let store = MetricsStore::new();
    // Observe values: 0.001, 0.02, 0.5, 2.0
    store.record_success("docx", 0.001, 100, 200, 1);
    store.record_success("docx", 0.02, 100, 200, 1);
    store.record_success("docx", 0.5, 100, 200, 1);
    store.record_success("docx", 2.0, 100, 200, 1);

    let output = store.render();
    // le=0.01: 0.001 fits => 1
    assert!(
        output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"0.01\"} 1"
        )
    );
    // le=0.05: 0.001, 0.02 fit => 2
    assert!(
        output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"0.05\"} 2"
        )
    );
    // le=0.5: 0.001, 0.02, 0.5 fit => 3
    assert!(
        output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"0.5\"} 3"
        )
    );
    // le=2.5: all 4 fit => 4
    assert!(
        output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"2.5\"} 4"
        )
    );
    // +Inf: all => 4
    assert!(
        output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"+Inf\"} 4"
        )
    );
}

#[test]
fn test_input_bytes_histogram() {
    let store = MetricsStore::new();
    store.record_success("xlsx", 0.1, 50_000, 100_000, 5);

    let output = store.render();
    // 50_000 bytes is between 10_240 and 102_400
    assert!(
        output.contains("office2pdf_conversion_input_bytes_bucket{format=\"xlsx\",le=\"10240\"} 0")
    );
    assert!(
        output
            .contains("office2pdf_conversion_input_bytes_bucket{format=\"xlsx\",le=\"102400\"} 1")
    );
}

#[test]
fn test_pages_histogram() {
    let store = MetricsStore::new();
    store.record_success("docx", 0.1, 100, 200, 7);

    let output = store.render();
    // 7 pages: le=5 -> 0, le=10 -> 1
    assert!(output.contains("office2pdf_conversion_pages_bucket{format=\"docx\",le=\"5\"} 0"));
    assert!(output.contains("office2pdf_conversion_pages_bucket{format=\"docx\",le=\"10\"} 1"));
}

#[test]
fn test_format_to_label() {
    use office2pdf::config::Format;
    assert_eq!(format_to_label(Format::Docx), "docx");
    assert_eq!(format_to_label(Format::Pptx), "pptx");
    assert_eq!(format_to_label(Format::Xlsx), "xlsx");
}

#[test]
fn test_render_has_all_help_lines() {
    let store = MetricsStore::new();
    store.record_success("docx", 0.1, 100, 200, 1);
    let output = store.render();

    assert!(output.contains("# HELP office2pdf_conversions_total"));
    assert!(output.contains("# HELP office2pdf_errors_total"));
    assert!(output.contains("# HELP office2pdf_conversion_duration_seconds"));
    assert!(output.contains("# HELP office2pdf_conversion_input_bytes"));
    assert!(output.contains("# HELP office2pdf_conversion_output_bytes"));
    assert!(output.contains("# HELP office2pdf_conversion_pages"));
    assert!(output.contains("# HELP office2pdf_active_conversions"));
}

#[test]
fn test_render_has_all_type_lines() {
    let store = MetricsStore::new();
    store.record_success("docx", 0.1, 100, 200, 1);
    let output = store.render();

    assert!(output.contains("# TYPE office2pdf_conversions_total counter"));
    assert!(output.contains("# TYPE office2pdf_errors_total counter"));
    assert!(output.contains("# TYPE office2pdf_conversion_duration_seconds histogram"));
    assert!(output.contains("# TYPE office2pdf_conversion_input_bytes histogram"));
    assert!(output.contains("# TYPE office2pdf_conversion_output_bytes histogram"));
    assert!(output.contains("# TYPE office2pdf_conversion_pages histogram"));
    assert!(output.contains("# TYPE office2pdf_active_conversions gauge"));
}

#[test]
fn test_error_types_tracked_separately() {
    let store = MetricsStore::new();
    store.record_failure("docx", "conversion");
    store.record_failure("docx", "conversion");
    store.record_failure("docx", "invalid_request");

    let output = store.render();
    assert!(
        output.contains("office2pdf_errors_total{format=\"docx\",error_type=\"conversion\"} 2")
    );
    assert!(
        output
            .contains("office2pdf_errors_total{format=\"docx\",error_type=\"invalid_request\"} 1")
    );
}

#[test]
fn test_output_bytes_histogram() {
    let store = MetricsStore::new();
    store.record_success("pptx", 0.1, 100, 5_000_000, 10);

    let output = store.render();
    // 5_000_000 between 1_048_576 and 10_485_760
    assert!(
        output.contains(
            "office2pdf_conversion_output_bytes_bucket{format=\"pptx\",le=\"1048576\"} 0"
        )
    );
    assert!(
        output.contains(
            "office2pdf_conversion_output_bytes_bucket{format=\"pptx\",le=\"10485760\"} 1"
        )
    );
}

#[test]
fn test_histogram_sum_accumulates() {
    let store = MetricsStore::new();
    store.record_success("docx", 1.5, 100, 200, 1);
    store.record_success("docx", 2.5, 100, 200, 1);

    let output = store.render();
    // Sum should be 4.0
    assert!(output.contains("office2pdf_conversion_duration_seconds_sum{format=\"docx\"} 4"));
    assert!(output.contains("office2pdf_conversion_duration_seconds_count{format=\"docx\"} 2"));
}
