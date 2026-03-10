//! HTTP server mode for office2pdf.
//!
//! Provides a REST API for document conversion via `office2pdf serve`.

use std::collections::HashMap;
use std::sync::Arc;

use anyhow::Result;
use office2pdf::config::{ConvertOptions, Format, PaperSize};

use crate::metrics::{self, MetricsStore};

/// Start the HTTP server on the given host and port.
pub fn start_server(host: &str, port: u16) -> Result<()> {
    let addr = format!("{host}:{port}");
    let server = tiny_http::Server::http(&addr)
        .map_err(|e| anyhow::anyhow!("failed to bind to {addr}: {e}"))?;

    let metrics = Arc::new(MetricsStore::new());

    eprintln!("office2pdf server listening on http://{addr}");
    eprintln!("Endpoints:");
    eprintln!("  POST /convert  - Convert a document to PDF");
    eprintln!("  GET  /health   - Health check");
    eprintln!("  GET  /formats  - List supported formats");
    eprintln!("  GET  /metrics  - Prometheus metrics");

    for mut request in server.incoming_requests() {
        let response = dispatch(&mut request, &metrics);
        let _ = request.respond(response);
    }

    Ok(())
}

type Response = tiny_http::Response<std::io::Cursor<Vec<u8>>>;

fn json_header() -> tiny_http::Header {
    tiny_http::Header::from_bytes("Content-Type", "application/json").unwrap()
}

fn pdf_header() -> tiny_http::Header {
    tiny_http::Header::from_bytes("Content-Type", "application/pdf").unwrap()
}

fn text_plain_header() -> tiny_http::Header {
    tiny_http::Header::from_bytes("Content-Type", "text/plain; version=0.0.4; charset=utf-8")
        .unwrap()
}

fn json_response(status: i32, body: &str) -> Response {
    tiny_http::Response::from_string(body)
        .with_header(json_header())
        .with_status_code(status)
}

fn dispatch(request: &mut tiny_http::Request, metrics: &MetricsStore) -> Response {
    let url = request.url().to_string();
    let path = url.split('?').next().unwrap_or(&url).to_string();
    let is_get = *request.method() == tiny_http::Method::Get;
    let is_post = *request.method() == tiny_http::Method::Post;

    if is_get && path == "/health" {
        handle_health()
    } else if is_get && path == "/formats" {
        handle_formats()
    } else if is_get && path == "/metrics" {
        handle_metrics(metrics)
    } else if is_post && path == "/convert" {
        handle_convert(request, &url, metrics)
    } else {
        json_response(404, r#"{"error":"not found"}"#)
    }
}

fn handle_health() -> Response {
    let version = env!("CARGO_PKG_VERSION");
    json_response(200, &format!(r#"{{"status":"ok","version":"{version}"}}"#))
}

fn handle_formats() -> Response {
    json_response(200, r#"{"formats":["docx","pptx","xlsx"]}"#)
}

fn handle_metrics(metrics: &MetricsStore) -> Response {
    let body = metrics.render();
    tiny_http::Response::from_string(body)
        .with_header(text_plain_header())
        .with_status_code(200)
}

fn handle_convert(request: &mut tiny_http::Request, url: &str, metrics: &MetricsStore) -> Response {
    metrics.start_conversion();
    let result = handle_convert_inner(request, url);
    metrics.end_conversion();

    match result {
        Ok(outcome) => {
            let format_label = metrics::format_to_label(outcome.format);
            if let Some(ref m) = outcome.metrics {
                metrics.record_success(
                    format_label,
                    m.total_duration.as_secs_f64(),
                    m.input_size_bytes,
                    m.output_size_bytes,
                    m.page_count,
                );
            } else {
                metrics.record_success(format_label, 0.0, 0, 0, 0);
            }
            tiny_http::Response::from_data(outcome.pdf)
                .with_header(pdf_header())
                .with_status_code(200)
        }
        Err(failure) => {
            metrics.record_failure(&failure.format_label, &failure.error_type);
            let msg = failure.message.replace('"', "\\\"");
            json_response(400, &format!(r#"{{"error":"{msg}"}}"#))
        }
    }
}

struct ConvertOutcome {
    pdf: Vec<u8>,
    format: Format,
    metrics: Option<office2pdf::error::ConvertMetrics>,
}

struct ConvertFailure {
    message: String,
    format_label: String,
    error_type: String,
}

fn handle_convert_inner(
    request: &mut tiny_http::Request,
    url: &str,
) -> std::result::Result<ConvertOutcome, ConvertFailure> {
    // Read body
    let mut body = Vec::new();
    request
        .as_reader()
        .read_to_end(&mut body)
        .map_err(|e| ConvertFailure {
            message: e.to_string(),
            format_label: "unknown".to_string(),
            error_type: "invalid_request".to_string(),
        })?;

    // Get content type header
    let content_type = request
        .headers()
        .iter()
        .find(|h| h.field.equiv("Content-Type"))
        .map(|h| h.value.as_str().to_string())
        .unwrap_or_default();

    // Parse multipart
    let boundary = extract_boundary(&content_type).ok_or_else(|| ConvertFailure {
        message: "missing or invalid Content-Type boundary".to_string(),
        format_label: "unknown".to_string(),
        error_type: "invalid_request".to_string(),
    })?;
    let file = extract_file_from_multipart(&body, &boundary).ok_or_else(|| ConvertFailure {
        message: "no file found in multipart body".to_string(),
        format_label: "unknown".to_string(),
        error_type: "invalid_request".to_string(),
    })?;

    // Parse query parameters
    let query = parse_query_string(url);

    // Detect format
    let format = if let Some(fmt) = query.get("format") {
        Format::from_extension(fmt).ok_or_else(|| ConvertFailure {
            message: format!("unsupported format: {fmt}"),
            format_label: "unknown".to_string(),
            error_type: "unsupported_format".to_string(),
        })?
    } else {
        detect_format_from_filename(&file.filename).ok_or_else(|| ConvertFailure {
            message: format!("cannot detect format from filename: {}", file.filename),
            format_label: "unknown".to_string(),
            error_type: "unsupported_format".to_string(),
        })?
    };

    let format_label = metrics::format_to_label(format).to_string();

    // Build options
    let mut options = ConvertOptions::default();
    if let Some(paper) = query.get("paper") {
        options.paper_size = Some(PaperSize::parse(paper).map_err(|e| ConvertFailure {
            message: e.to_string(),
            format_label: format_label.clone(),
            error_type: "invalid_request".to_string(),
        })?);
    }
    if let Some(landscape) = query.get("landscape")
        && (landscape == "true" || landscape == "1")
    {
        options.landscape = Some(true);
    }

    // Convert
    let result =
        office2pdf::convert_bytes(&file.data, format, &options).map_err(|e| ConvertFailure {
            message: format!("conversion failed: {e}"),
            format_label,
            error_type: "conversion".to_string(),
        })?;

    Ok(ConvertOutcome {
        pdf: result.pdf,
        format,
        metrics: result.metrics,
    })
}

// --- Multipart parsing helpers ---

struct MultipartFile {
    filename: String,
    data: Vec<u8>,
}

fn extract_boundary(content_type: &str) -> Option<String> {
    content_type.split(';').find_map(|part| {
        let part = part.trim();
        part.strip_prefix("boundary=")
            .map(|b| b.trim_matches('"').to_string())
    })
}

fn extract_file_from_multipart(body: &[u8], boundary: &str) -> Option<MultipartFile> {
    let delim = format!("--{boundary}");
    let delim_bytes = delim.as_bytes();

    // Find the first delimiter
    let first_pos = find_bytes(body, delim_bytes)?;
    let after_delim = first_pos + delim_bytes.len();

    // Skip \r\n after delimiter
    let start = if body.get(after_delim..after_delim + 2) == Some(b"\r\n") {
        after_delim + 2
    } else {
        after_delim
    };

    // Find \r\n\r\n (headers/body separator)
    let header_end = find_bytes(&body[start..], b"\r\n\r\n")?;
    let headers = std::str::from_utf8(&body[start..start + header_end]).ok()?;
    let data_start = start + header_end + 4;

    // Find the next delimiter to determine data end
    let next_delim_pos = find_bytes(&body[data_start..], delim_bytes)?;
    // Data ends before \r\n that precedes the next delimiter
    let data_end = if next_delim_pos >= 2
        && body[data_start + next_delim_pos - 2..data_start + next_delim_pos] == *b"\r\n"
    {
        data_start + next_delim_pos - 2
    } else {
        data_start + next_delim_pos
    };

    let filename = extract_filename_from_headers(headers)?;

    Some(MultipartFile {
        filename,
        data: body[data_start..data_end].to_vec(),
    })
}

fn find_bytes(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack.windows(needle.len()).position(|w| w == needle)
}

fn extract_filename_from_headers(headers: &str) -> Option<String> {
    let lower = headers.to_ascii_lowercase();
    let idx = lower.find("filename=\"")?;
    let start = idx + "filename=\"".len();
    let rest = &headers[start..];
    let end = rest.find('"')?;
    Some(rest[..end].to_string())
}

fn detect_format_from_filename(filename: &str) -> Option<Format> {
    let ext = filename.rsplit('.').next()?;
    Format::from_extension(ext)
}

fn parse_query_string(url: &str) -> HashMap<String, String> {
    let mut params = HashMap::new();
    if let Some(query) = url.split('?').nth(1) {
        for pair in query.split('&') {
            if let Some((key, value)) = pair.split_once('=') {
                params.insert(key.to_string(), value.to_string());
            }
        }
    }
    params
}

#[cfg(test)]
#[path = "server_tests.rs"]
mod tests;
