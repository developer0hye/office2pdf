//! Prometheus-compatible metrics for the office2pdf server.
//!
//! Provides an in-memory metrics store that tracks conversion counters,
//! histograms (duration, bytes, pages), and an active-conversions gauge.
//! The `/metrics` endpoint renders these in Prometheus exposition format.

use std::collections::BTreeMap;
use std::fmt::Write;
use std::sync::Mutex;
use std::sync::atomic::{AtomicI64, Ordering};

/// Pre-defined histogram buckets for conversion duration (seconds).
const DURATION_BUCKETS: &[f64] = &[0.01, 0.05, 0.1, 0.25, 0.5, 1.0, 2.5, 5.0, 10.0, 30.0, 60.0];

/// Pre-defined histogram buckets for file sizes (bytes).
const BYTES_BUCKETS: &[f64] = &[
    1024.0,
    10_240.0,
    102_400.0,
    1_048_576.0,
    10_485_760.0,
    104_857_600.0,
    1_073_741_824.0,
];

/// Pre-defined histogram buckets for page counts.
const PAGES_BUCKETS: &[f64] = &[1.0, 5.0, 10.0, 25.0, 50.0, 100.0, 500.0, 1000.0];

/// A single histogram with pre-defined bucket boundaries.
struct Histogram {
    buckets: &'static [f64],
    /// Cumulative count of observations <= each bucket boundary.
    counts: Vec<u64>,
    sum: f64,
    count: u64,
}

impl Histogram {
    fn new(buckets: &'static [f64]) -> Self {
        Self {
            buckets,
            counts: vec![0; buckets.len()],
            sum: 0.0,
            count: 0,
        }
    }

    fn observe(&mut self, value: f64) {
        for (i, bound) in self.buckets.iter().enumerate() {
            if value <= *bound {
                self.counts[i] += 1;
            }
        }
        self.sum += value;
        self.count += 1;
    }
}

/// Thread-safe metrics store for Prometheus-compatible monitoring.
pub struct MetricsStore {
    /// Conversion counters: (format, status) -> count.
    conversions: Mutex<BTreeMap<(String, String), u64>>,
    /// Error counters: (format, error_type) -> count.
    errors: Mutex<BTreeMap<(String, String), u64>>,
    /// Conversion duration histogram by format.
    duration: Mutex<BTreeMap<String, Histogram>>,
    /// Input size histogram by format.
    input_bytes: Mutex<BTreeMap<String, Histogram>>,
    /// Output size histogram by format.
    output_bytes: Mutex<BTreeMap<String, Histogram>>,
    /// Page count histogram by format.
    pages: Mutex<BTreeMap<String, Histogram>>,
    /// Currently active (in-progress) conversions.
    active: AtomicI64,
}

impl MetricsStore {
    /// Create an empty metrics store.
    pub fn new() -> Self {
        Self {
            conversions: Mutex::new(BTreeMap::new()),
            errors: Mutex::new(BTreeMap::new()),
            duration: Mutex::new(BTreeMap::new()),
            input_bytes: Mutex::new(BTreeMap::new()),
            output_bytes: Mutex::new(BTreeMap::new()),
            pages: Mutex::new(BTreeMap::new()),
            active: AtomicI64::new(0),
        }
    }

    /// Increment the active-conversions gauge (call before conversion starts).
    pub fn start_conversion(&self) {
        self.active.fetch_add(1, Ordering::Relaxed);
    }

    /// Decrement the active-conversions gauge (call after conversion finishes).
    pub fn end_conversion(&self) {
        self.active.fetch_sub(1, Ordering::Relaxed);
    }

    /// Record a successful conversion with its metrics.
    pub fn record_success(
        &self,
        format: &str,
        duration_secs: f64,
        input_size: u64,
        output_size: u64,
        page_count: u32,
    ) {
        *self
            .conversions
            .lock()
            .unwrap()
            .entry((format.to_string(), "success".to_string()))
            .or_insert(0) += 1;

        self.duration
            .lock()
            .unwrap()
            .entry(format.to_string())
            .or_insert_with(|| Histogram::new(DURATION_BUCKETS))
            .observe(duration_secs);

        self.input_bytes
            .lock()
            .unwrap()
            .entry(format.to_string())
            .or_insert_with(|| Histogram::new(BYTES_BUCKETS))
            .observe(input_size as f64);

        self.output_bytes
            .lock()
            .unwrap()
            .entry(format.to_string())
            .or_insert_with(|| Histogram::new(BYTES_BUCKETS))
            .observe(output_size as f64);

        self.pages
            .lock()
            .unwrap()
            .entry(format.to_string())
            .or_insert_with(|| Histogram::new(PAGES_BUCKETS))
            .observe(page_count as f64);
    }

    /// Record a failed conversion.
    pub fn record_failure(&self, format: &str, error_type: &str) {
        *self
            .conversions
            .lock()
            .unwrap()
            .entry((format.to_string(), "failure".to_string()))
            .or_insert(0) += 1;

        *self
            .errors
            .lock()
            .unwrap()
            .entry((format.to_string(), error_type.to_string()))
            .or_insert(0) += 1;
    }

    /// Render all metrics in Prometheus exposition text format.
    pub fn render(&self) -> String {
        let mut out = String::new();

        self.render_conversions(&mut out);
        self.render_errors(&mut out);
        self.render_histogram_metric(
            &mut out,
            "office2pdf_conversion_duration_seconds",
            "Duration of document conversion in seconds",
            &self.duration,
        );
        self.render_histogram_metric(
            &mut out,
            "office2pdf_conversion_input_bytes",
            "Size of input documents in bytes",
            &self.input_bytes,
        );
        self.render_histogram_metric(
            &mut out,
            "office2pdf_conversion_output_bytes",
            "Size of output PDFs in bytes",
            &self.output_bytes,
        );
        self.render_histogram_metric(
            &mut out,
            "office2pdf_conversion_pages",
            "Number of pages in output PDFs",
            &self.pages,
        );
        self.render_active(&mut out);

        out
    }

    fn render_conversions(&self, out: &mut String) {
        let map = self.conversions.lock().unwrap();
        writeln!(
            out,
            "# HELP office2pdf_conversions_total Total number of document conversions"
        )
        .unwrap();
        writeln!(out, "# TYPE office2pdf_conversions_total counter").unwrap();
        for ((format, status), count) in map.iter() {
            writeln!(
                out,
                "office2pdf_conversions_total{{format=\"{format}\",status=\"{status}\"}} {count}"
            )
            .unwrap();
        }
    }

    fn render_errors(&self, out: &mut String) {
        let map = self.errors.lock().unwrap();
        writeln!(
            out,
            "# HELP office2pdf_errors_total Total number of conversion errors"
        )
        .unwrap();
        writeln!(out, "# TYPE office2pdf_errors_total counter").unwrap();
        for ((format, error_type), count) in map.iter() {
            writeln!(
                out,
                "office2pdf_errors_total{{format=\"{format}\",error_type=\"{error_type}\"}} {count}"
            )
            .unwrap();
        }
    }

    fn render_histogram_metric(
        &self,
        out: &mut String,
        name: &str,
        help: &str,
        data: &Mutex<BTreeMap<String, Histogram>>,
    ) {
        let map = data.lock().unwrap();
        writeln!(out, "# HELP {name} {help}").unwrap();
        writeln!(out, "# TYPE {name} histogram").unwrap();
        for (format, hist) in map.iter() {
            for (i, bound) in hist.buckets.iter().enumerate() {
                writeln!(
                    out,
                    "{name}_bucket{{format=\"{format}\",le=\"{bound}\"}} {}",
                    hist.counts[i]
                )
                .unwrap();
            }
            writeln!(
                out,
                "{name}_bucket{{format=\"{format}\",le=\"+Inf\"}} {}",
                hist.count
            )
            .unwrap();
            writeln!(out, "{name}_sum{{format=\"{format}\"}} {}", hist.sum).unwrap();
            writeln!(out, "{name}_count{{format=\"{format}\"}} {}", hist.count).unwrap();
        }
    }

    fn render_active(&self, out: &mut String) {
        let val = self.active.load(Ordering::Relaxed);
        writeln!(
            out,
            "# HELP office2pdf_active_conversions Number of currently active conversions"
        )
        .unwrap();
        writeln!(out, "# TYPE office2pdf_active_conversions gauge").unwrap();
        writeln!(out, "office2pdf_active_conversions {val}").unwrap();
    }
}

/// Map a `Format` enum variant to its lowercase label string.
pub fn format_to_label(format: office2pdf::config::Format) -> &'static str {
    match format {
        office2pdf::config::Format::Docx => "docx",
        office2pdf::config::Format::Pptx => "pptx",
        office2pdf::config::Format::Xlsx => "xlsx",
    }
}

#[cfg(test)]
mod tests {
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
        assert!(
            output.contains("office2pdf_conversions_total{format=\"docx\",status=\"success\"} 2")
        );
    }

    #[test]
    fn test_record_failure_increments_counters() {
        let store = MetricsStore::new();
        store.record_failure("pptx", "conversion");

        let output = store.render();
        assert!(
            output.contains("office2pdf_conversions_total{format=\"pptx\",status=\"failure\"} 1")
        );
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
        assert!(output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"0.05\"} 1"
        ));
        // Should NOT be in le=0.01 bucket
        assert!(output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"0.01\"} 0"
        ));
        // Should be in +Inf bucket
        assert!(output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"+Inf\"} 1"
        ));
        // Sum and count
        assert!(
            output.contains("office2pdf_conversion_duration_seconds_sum{format=\"docx\"} 0.05")
        );
        assert!(output.contains("office2pdf_conversion_duration_seconds_count{format=\"docx\"} 1"));
    }

    #[test]
    fn test_multiple_formats_tracked_separately() {
        let store = MetricsStore::new();
        store.record_success("docx", 0.1, 100, 200, 1);
        store.record_success("xlsx", 0.2, 300, 400, 2);
        store.record_failure("pptx", "conversion");

        let output = store.render();
        assert!(
            output.contains("office2pdf_conversions_total{format=\"docx\",status=\"success\"} 1")
        );
        assert!(
            output.contains("office2pdf_conversions_total{format=\"xlsx\",status=\"success\"} 1")
        );
        assert!(
            output.contains("office2pdf_conversions_total{format=\"pptx\",status=\"failure\"} 1")
        );
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
        assert!(output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"0.01\"} 1"
        ));
        // le=0.05: 0.001, 0.02 fit => 2
        assert!(output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"0.05\"} 2"
        ));
        // le=0.5: 0.001, 0.02, 0.5 fit => 3
        assert!(output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"0.5\"} 3"
        ));
        // le=2.5: all 4 fit => 4
        assert!(output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"2.5\"} 4"
        ));
        // +Inf: all => 4
        assert!(output.contains(
            "office2pdf_conversion_duration_seconds_bucket{format=\"docx\",le=\"+Inf\"} 4"
        ));
    }

    #[test]
    fn test_input_bytes_histogram() {
        let store = MetricsStore::new();
        store.record_success("xlsx", 0.1, 50_000, 100_000, 5);

        let output = store.render();
        // 50_000 bytes is between 10_240 and 102_400
        assert!(
            output.contains(
                "office2pdf_conversion_input_bytes_bucket{format=\"xlsx\",le=\"10240\"} 0"
            )
        );
        assert!(
            output.contains(
                "office2pdf_conversion_input_bytes_bucket{format=\"xlsx\",le=\"102400\"} 1"
            )
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
            output.contains(
                "office2pdf_errors_total{format=\"docx\",error_type=\"invalid_request\"} 1"
            )
        );
    }

    #[test]
    fn test_output_bytes_histogram() {
        let store = MetricsStore::new();
        store.record_success("pptx", 0.1, 100, 5_000_000, 10);

        let output = store.render();
        // 5_000_000 between 1_048_576 and 10_485_760
        assert!(output.contains(
            "office2pdf_conversion_output_bytes_bucket{format=\"pptx\",le=\"1048576\"} 0"
        ));
        assert!(output.contains(
            "office2pdf_conversion_output_bytes_bucket{format=\"pptx\",le=\"10485760\"} 1"
        ));
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
}
