//! HTTP server mode for office2pdf.
//!
//! Provides a REST API for document conversion via `office2pdf serve`.

use std::collections::HashMap;

use anyhow::Result;
use office2pdf::config::{ConvertOptions, Format, PaperSize};

/// Start the HTTP server on the given host and port.
pub fn start_server(host: &str, port: u16) -> Result<()> {
    let addr = format!("{host}:{port}");
    let server = tiny_http::Server::http(&addr)
        .map_err(|e| anyhow::anyhow!("failed to bind to {addr}: {e}"))?;

    eprintln!("office2pdf server listening on http://{addr}");
    eprintln!("Endpoints:");
    eprintln!("  POST /convert  - Convert a document to PDF");
    eprintln!("  GET  /health   - Health check");
    eprintln!("  GET  /formats  - List supported formats");

    for mut request in server.incoming_requests() {
        let response = dispatch(&mut request);
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

fn json_response(status: i32, body: &str) -> Response {
    tiny_http::Response::from_string(body)
        .with_header(json_header())
        .with_status_code(status)
}

fn dispatch(request: &mut tiny_http::Request) -> Response {
    let url = request.url().to_string();
    let path = url.split('?').next().unwrap_or(&url).to_string();
    let is_get = *request.method() == tiny_http::Method::Get;
    let is_post = *request.method() == tiny_http::Method::Post;

    if is_get && path == "/health" {
        handle_health()
    } else if is_get && path == "/formats" {
        handle_formats()
    } else if is_post && path == "/convert" {
        handle_convert(request, &url)
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

fn handle_convert(request: &mut tiny_http::Request, url: &str) -> Response {
    match handle_convert_inner(request, url) {
        Ok(pdf_bytes) => tiny_http::Response::from_data(pdf_bytes)
            .with_header(pdf_header())
            .with_status_code(200),
        Err(e) => {
            let msg = e.to_string().replace('"', "\\\"");
            json_response(400, &format!(r#"{{"error":"{msg}"}}"#))
        }
    }
}

fn handle_convert_inner(request: &mut tiny_http::Request, url: &str) -> Result<Vec<u8>> {
    // Read body
    let mut body = Vec::new();
    request.as_reader().read_to_end(&mut body)?;

    // Get content type header
    let content_type = request
        .headers()
        .iter()
        .find(|h| h.field.equiv("Content-Type"))
        .map(|h| h.value.as_str().to_string())
        .unwrap_or_default();

    // Parse multipart
    let boundary = extract_boundary(&content_type)
        .ok_or_else(|| anyhow::anyhow!("missing or invalid Content-Type boundary"))?;
    let file = extract_file_from_multipart(&body, &boundary)
        .ok_or_else(|| anyhow::anyhow!("no file found in multipart body"))?;

    // Parse query parameters
    let query = parse_query_string(url);

    // Detect format
    let format = if let Some(fmt) = query.get("format") {
        Format::from_extension(fmt).ok_or_else(|| anyhow::anyhow!("unsupported format: {fmt}"))?
    } else {
        detect_format_from_filename(&file.filename).ok_or_else(|| {
            anyhow::anyhow!("cannot detect format from filename: {}", file.filename)
        })?
    };

    // Build options
    let mut options = ConvertOptions::default();
    if let Some(paper) = query.get("paper") {
        options.paper_size = Some(PaperSize::parse(paper).map_err(|e| anyhow::anyhow!("{e}"))?);
    }
    if let Some(landscape) = query.get("landscape")
        && (landscape == "true" || landscape == "1")
    {
        options.landscape = Some(true);
    }

    // Convert
    let result = office2pdf::convert_bytes(&file.data, format, &options)
        .map_err(|e| anyhow::anyhow!("conversion failed: {e}"))?;

    Ok(result.pdf)
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
mod tests {
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
    fn start_test_server(n: usize) -> (std::thread::JoinHandle<()>, u16) {
        let server = tiny_http::Server::http("127.0.0.1:0").unwrap();
        let port = match server.server_addr() {
            tiny_http::ListenAddr::IP(addr) => addr.port(),
            _ => panic!("expected IP address"),
        };

        let handle = std::thread::spawn(move || {
            for _ in 0..n {
                if let Ok(mut request) = server.recv() {
                    let response = dispatch(&mut request);
                    let _ = request.respond(response);
                }
            }
        });

        (handle, port)
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
        let (handle, port) = start_test_server(1);
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
        let (handle, port) = start_test_server(1);
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
        let (handle, port) = start_test_server(1);
        let addr = format!("127.0.0.1:{port}");

        let resp = send_request(&addr, "GET", "/nonexistent", &[], &[]);

        assert_eq!(resp.status_code, 404);
        let body = resp.body_str();
        assert!(body.contains("\"error\""));

        handle.join().unwrap();
    }

    #[test]
    fn test_convert_docx_to_pdf() {
        let (handle, port) = start_test_server(1);
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
        let (handle, port) = start_test_server(1);
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
        let (handle, port) = start_test_server(1);
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
}
