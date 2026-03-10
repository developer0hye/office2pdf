use super::*;

// --- Unit tests for helper functions ---

#[test]
fn test_extract_boundary() {
    assert_eq!(
        extract_boundary("multipart/form-data; boundary=abc123"),
        Some("abc123".to_string())
    );
    assert_eq!(
        extract_boundary("multipart/form-data; boundary=\"abc123\""),
        Some("abc123".to_string())
    );
    assert_eq!(extract_boundary("application/json"), None);
    assert_eq!(extract_boundary(""), None);
}

#[test]
fn test_extract_filename_from_headers() {
    assert_eq!(
        extract_filename_from_headers(
            "Content-Disposition: form-data; name=\"file\"; filename=\"report.docx\""
        ),
        Some("report.docx".to_string())
    );
    assert_eq!(
        extract_filename_from_headers(
            "content-disposition: form-data; name=\"file\"; filename=\"test.pptx\""
        ),
        Some("test.pptx".to_string())
    );
    assert_eq!(
        extract_filename_from_headers("Content-Type: application/octet-stream"),
        None
    );
}

#[test]
fn test_detect_format_from_filename() {
    assert_eq!(
        detect_format_from_filename("report.docx"),
        Some(Format::Docx)
    );
    assert_eq!(
        detect_format_from_filename("slides.pptx"),
        Some(Format::Pptx)
    );
    assert_eq!(detect_format_from_filename("data.xlsx"), Some(Format::Xlsx));
    assert_eq!(detect_format_from_filename("README.md"), None);
    assert_eq!(detect_format_from_filename("noext"), None);
}

#[test]
fn test_parse_query_string() {
    let params = parse_query_string("/convert?format=docx&paper=a4");
    assert_eq!(params.get("format").map(|s| s.as_str()), Some("docx"));
    assert_eq!(params.get("paper").map(|s| s.as_str()), Some("a4"));

    let params = parse_query_string("/convert");
    assert!(params.is_empty());
}

#[test]
fn test_extract_file_from_multipart() {
    let boundary = "TESTBOUNDARY";
    let body = build_multipart_body(b"hello world", "test.docx", boundary);
    let file = extract_file_from_multipart(&body, boundary).unwrap();
    assert_eq!(file.filename, "test.docx");
    assert_eq!(file.data, b"hello world");
}

#[test]
fn test_extract_file_from_multipart_binary() {
    let boundary = "BINBOUNDARY";
    let data: Vec<u8> = (0..=255).collect();
    let body = build_multipart_body(&data, "binary.bin", boundary);
    let file = extract_file_from_multipart(&body, boundary).unwrap();
    assert_eq!(file.filename, "binary.bin");
    assert_eq!(file.data, data);
}

fn build_multipart_body(file_data: &[u8], filename: &str, boundary: &str) -> Vec<u8> {
    let mut body = Vec::new();
    body.extend_from_slice(format!("--{boundary}\r\n").as_bytes());
    body.extend_from_slice(
        format!("Content-Disposition: form-data; name=\"file\"; filename=\"{filename}\"\r\n")
            .as_bytes(),
    );
    body.extend_from_slice(b"Content-Type: application/octet-stream\r\n");
    body.extend_from_slice(b"\r\n");
    body.extend_from_slice(file_data);
    body.extend_from_slice(format!("\r\n--{boundary}--\r\n").as_bytes());
    body
}

// --- Integration tests ---

fn make_test_docx() -> Vec<u8> {
    use std::io::Cursor;
    let docx = docx_rs::Docx::new().add_paragraph(
        docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Hello server")),
    );
    let mut buf = Cursor::new(Vec::new());
    docx.build().pack(&mut buf).unwrap();
    buf.into_inner()
}

/// Start a server on an ephemeral port, handle `n` requests, then return.
fn start_test_server(n: usize) -> (std::thread::JoinHandle<()>, u16, Arc<MetricsStore>) {
    let server = tiny_http::Server::http("127.0.0.1:0").unwrap();
    let port = match server.server_addr() {
        tiny_http::ListenAddr::IP(addr) => addr.port(),
        _ => panic!("expected IP address"),
    };

    let metrics = Arc::new(MetricsStore::new());
    let metrics_clone = Arc::clone(&metrics);

    let handle = std::thread::spawn(move || {
        for _ in 0..n {
            if let Ok(mut request) = server.recv() {
                let response = dispatch(&mut request, &metrics_clone);
                let _ = request.respond(response);
            }
        }
    });

    (handle, port, metrics)
}

struct HttpResponse {
    status_code: u16,
    #[allow(dead_code)]
    headers: HashMap<String, String>,
    body: Vec<u8>,
}

impl HttpResponse {
    fn body_str(&self) -> String {
        String::from_utf8_lossy(&self.body).to_string()
    }

    fn content_type(&self) -> Option<&str> {
        self.headers.get("content-type").map(|s| s.as_str())
    }
}

fn send_request(
    addr: &str,
    method: &str,
    path: &str,
    extra_headers: &[(&str, &str)],
    body: &[u8],
) -> HttpResponse {
    use std::io::{BufRead, BufReader, Read, Write};
    use std::net::TcpStream;
    use std::time::Duration;

    let mut stream = TcpStream::connect(addr).unwrap();
    stream
        .set_read_timeout(Some(Duration::from_secs(60)))
        .unwrap();

    // Write request
    write!(stream, "{method} {path} HTTP/1.1\r\n").unwrap();
    write!(stream, "Host: {addr}\r\n").unwrap();
    write!(stream, "Connection: close\r\n").unwrap();
    if !body.is_empty() {
        write!(stream, "Content-Length: {}\r\n", body.len()).unwrap();
    }
    for (key, value) in extra_headers {
        write!(stream, "{key}: {value}\r\n").unwrap();
    }
    write!(stream, "\r\n").unwrap();
    if !body.is_empty() {
        stream.write_all(body).unwrap();
    }
    stream.flush().unwrap();

    // Read response
    let mut reader = BufReader::new(&stream);

    // Status line
    let mut status_line = String::new();
    reader.read_line(&mut status_line).unwrap();
    let status_code: u16 = status_line
        .split(' ')
        .nth(1)
        .unwrap()
        .trim()
        .parse()
        .unwrap();

    // Headers
    let mut resp_headers = HashMap::new();
    let mut content_length = 0usize;
    loop {
        let mut line = String::new();
        reader.read_line(&mut line).unwrap();
        let trimmed = line.trim();
        if trimmed.is_empty() {
            break;
        }
        if let Some((key, value)) = trimmed.split_once(':') {
            let key = key.trim().to_ascii_lowercase();
            let value = value.trim().to_string();
            if key == "content-length" {
                content_length = value.parse().unwrap_or(0);
            }
            resp_headers.insert(key, value);
        }
    }

    // Body
    let mut resp_body = vec![0u8; content_length];
    if content_length > 0 {
        reader.read_exact(&mut resp_body).unwrap();
    }

    HttpResponse {
        status_code,
        headers: resp_headers,
        body: resp_body,
    }
}

#[test]
fn test_health_endpoint() {
    let (handle, port, _metrics) = start_test_server(1);
    let addr = format!("127.0.0.1:{port}");

    let resp = send_request(&addr, "GET", "/health", &[], &[]);

    assert_eq!(resp.status_code, 200);
    assert!(resp.content_type().unwrap().contains("application/json"));
    let body = resp.body_str();
    assert!(body.contains("\"status\":\"ok\""));
    assert!(body.contains("\"version\""));

    handle.join().unwrap();
}

#[test]
fn test_formats_endpoint() {
    let (handle, port, _metrics) = start_test_server(1);
    let addr = format!("127.0.0.1:{port}");

    let resp = send_request(&addr, "GET", "/formats", &[], &[]);

    assert_eq!(resp.status_code, 200);
    let body = resp.body_str();
    assert!(body.contains("\"docx\""));
    assert!(body.contains("\"pptx\""));
    assert!(body.contains("\"xlsx\""));

    handle.join().unwrap();
}

#[test]
fn test_not_found_endpoint() {
    let (handle, port, _metrics) = start_test_server(1);
    let addr = format!("127.0.0.1:{port}");

    let resp = send_request(&addr, "GET", "/nonexistent", &[], &[]);

    assert_eq!(resp.status_code, 404);
    let body = resp.body_str();
    assert!(body.contains("\"error\""));

    handle.join().unwrap();
}

#[test]
fn test_convert_docx_to_pdf() {
    let (handle, port, _metrics) = start_test_server(1);
    let addr = format!("127.0.0.1:{port}");

    let docx_data = make_test_docx();
    let boundary = "TestBoundary12345";
    let multipart_body = build_multipart_body(&docx_data, "test.docx", boundary);
    let content_type = format!("multipart/form-data; boundary={boundary}");

    let resp = send_request(
        &addr,
        "POST",
        "/convert",
        &[("Content-Type", &content_type)],
        &multipart_body,
    );

    assert_eq!(resp.status_code, 200);
    assert!(resp.content_type().unwrap().contains("application/pdf"));
    assert!(
        resp.body.starts_with(b"%PDF"),
        "response should be a valid PDF"
    );

    handle.join().unwrap();
}

#[test]
fn test_convert_invalid_format_error() {
    let (handle, port, _metrics) = start_test_server(1);
    let addr = format!("127.0.0.1:{port}");

    let boundary = "TestBoundary67890";
    let multipart_body = build_multipart_body(b"not a document", "test.txt", boundary);
    let content_type = format!("multipart/form-data; boundary={boundary}");

    let resp = send_request(
        &addr,
        "POST",
        "/convert",
        &[("Content-Type", &content_type)],
        &multipart_body,
    );

    assert_eq!(resp.status_code, 400);
    let body = resp.body_str();
    assert!(body.contains("\"error\""));

    handle.join().unwrap();
}

#[test]
fn test_convert_with_format_override() {
    let (handle, port, _metrics) = start_test_server(1);
    let addr = format!("127.0.0.1:{port}");

    let docx_data = make_test_docx();
    let boundary = "FormatOverride";
    let multipart_body = build_multipart_body(&docx_data, "document", boundary);
    let content_type = format!("multipart/form-data; boundary={boundary}");

    let resp = send_request(
        &addr,
        "POST",
        "/convert?format=docx",
        &[("Content-Type", &content_type)],
        &multipart_body,
    );

    assert_eq!(resp.status_code, 200);
    assert!(
        resp.body.starts_with(b"%PDF"),
        "response should be a valid PDF"
    );

    handle.join().unwrap();
}

// --- Metrics endpoint tests ---

#[test]
fn test_metrics_endpoint_returns_200() {
    let (handle, port, _metrics) = start_test_server(1);
    let addr = format!("127.0.0.1:{port}");

    let resp = send_request(&addr, "GET", "/metrics", &[], &[]);

    assert_eq!(resp.status_code, 200);
    assert!(
        resp.content_type().unwrap().contains("text/plain"),
        "metrics endpoint should return text/plain"
    );
    let body = resp.body_str();
    assert!(body.contains("# HELP office2pdf_conversions_total"));
    assert!(body.contains("# TYPE office2pdf_active_conversions gauge"));
    assert!(body.contains("office2pdf_active_conversions 0"));

    handle.join().unwrap();
}

#[test]
fn test_metrics_after_successful_conversion() {
    // 2 requests: convert + metrics check
    let (handle, port, _metrics) = start_test_server(2);
    let addr = format!("127.0.0.1:{port}");

    // Step 1: Convert a DOCX file
    let docx_data = make_test_docx();
    let boundary = "MetricsTestBoundary";
    let multipart_body = build_multipart_body(&docx_data, "test.docx", boundary);
    let content_type = format!("multipart/form-data; boundary={boundary}");

    let convert_resp = send_request(
        &addr,
        "POST",
        "/convert",
        &[("Content-Type", &content_type)],
        &multipart_body,
    );
    assert_eq!(convert_resp.status_code, 200);

    // Step 2: Check metrics
    let metrics_resp = send_request(&addr, "GET", "/metrics", &[], &[]);
    assert_eq!(metrics_resp.status_code, 200);
    let body = metrics_resp.body_str();

    // Should show 1 successful docx conversion
    assert!(
        body.contains("office2pdf_conversions_total{format=\"docx\",status=\"success\"} 1"),
        "should track successful conversion: {body}"
    );
    // Should have duration histogram data
    assert!(
        body.contains("office2pdf_conversion_duration_seconds_count{format=\"docx\"} 1"),
        "should track duration histogram: {body}"
    );
    // Active conversions should be 0 (conversion finished)
    assert!(
        body.contains("office2pdf_active_conversions 0"),
        "active conversions should be 0 after conversion: {body}"
    );

    handle.join().unwrap();
}

#[test]
fn test_metrics_after_failed_conversion() {
    // 2 requests: failed convert + metrics check
    let (handle, port, _metrics) = start_test_server(2);
    let addr = format!("127.0.0.1:{port}");

    // Step 1: Try to convert an invalid file
    let boundary = "FailTestBoundary";
    let multipart_body = build_multipart_body(b"not valid", "test.txt", boundary);
    let content_type = format!("multipart/form-data; boundary={boundary}");

    let convert_resp = send_request(
        &addr,
        "POST",
        "/convert",
        &[("Content-Type", &content_type)],
        &multipart_body,
    );
    assert_eq!(convert_resp.status_code, 400);

    // Step 2: Check metrics
    let metrics_resp = send_request(&addr, "GET", "/metrics", &[], &[]);
    let body = metrics_resp.body_str();

    // Should show a failure with unknown format
    assert!(
        body.contains("office2pdf_conversions_total{format=\"unknown\",status=\"failure\"} 1"),
        "should track failed conversion: {body}"
    );
    assert!(
        body.contains(
            "office2pdf_errors_total{format=\"unknown\",error_type=\"unsupported_format\"} 1"
        ),
        "should track error type: {body}"
    );

    handle.join().unwrap();
}

#[test]
fn test_metrics_multiple_conversions() {
    // 3 requests: 2 converts + metrics check
    let (handle, port, _metrics) = start_test_server(3);
    let addr = format!("127.0.0.1:{port}");

    let docx_data = make_test_docx();
    let boundary = "MultiMetrics";
    let multipart_body = build_multipart_body(&docx_data, "test.docx", boundary);
    let content_type = format!("multipart/form-data; boundary={boundary}");

    // Two successful conversions
    let resp1 = send_request(
        &addr,
        "POST",
        "/convert",
        &[("Content-Type", &content_type)],
        &multipart_body,
    );
    assert_eq!(resp1.status_code, 200);

    let resp2 = send_request(
        &addr,
        "POST",
        "/convert",
        &[("Content-Type", &content_type)],
        &multipart_body,
    );
    assert_eq!(resp2.status_code, 200);

    // Check metrics
    let metrics_resp = send_request(&addr, "GET", "/metrics", &[], &[]);
    let body = metrics_resp.body_str();

    assert!(
        body.contains("office2pdf_conversions_total{format=\"docx\",status=\"success\"} 2"),
        "should show 2 successful conversions: {body}"
    );
    assert!(
        body.contains("office2pdf_conversion_duration_seconds_count{format=\"docx\"} 2"),
        "duration histogram should have count 2: {body}"
    );

    handle.join().unwrap();
}
