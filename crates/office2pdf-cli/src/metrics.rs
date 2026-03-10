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
#[path = "metrics_tests.rs"]
mod tests;
