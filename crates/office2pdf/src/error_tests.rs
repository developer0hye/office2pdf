use super::*;

#[test]
fn test_unsupported_element_display() {
    let w = ConvertWarning::UnsupportedElement {
        format: "DOCX".to_string(),
        element: "OLE object".to_string(),
    };
    assert_eq!(w.to_string(), "[DOCX] unsupported element: OLE object");
}

#[test]
fn test_partial_element_display() {
    let w = ConvertWarning::PartialElement {
        format: "PPTX".to_string(),
        element: "scheme color".to_string(),
        detail: "tint modifier ignored".to_string(),
    };
    assert_eq!(
        w.to_string(),
        "[PPTX] partial rendering of scheme color: tint modifier ignored"
    );
}

#[test]
fn test_fallback_used_display() {
    let w = ConvertWarning::FallbackUsed {
        format: "DOCX".to_string(),
        from: "chart".to_string(),
        to: "data table".to_string(),
    };
    assert_eq!(
        w.to_string(),
        "[DOCX] fallback: chart rendered as data table"
    );
}

#[test]
fn test_parse_skipped_display() {
    let w = ConvertWarning::ParseSkipped {
        format: "PPTX".to_string(),
        reason: "slide 3 failed to parse: missing XML".to_string(),
    };
    assert_eq!(
        w.to_string(),
        "[PPTX] skipped: slide 3 failed to parse: missing XML"
    );
}

#[test]
fn test_warning_format_accessor() {
    let w = ConvertWarning::FallbackUsed {
        format: "XLSX".to_string(),
        from: "chart".to_string(),
        to: "data table".to_string(),
    };
    assert_eq!(w.format(), "XLSX");
}

#[test]
fn test_warning_clone_and_eq() {
    let w = ConvertWarning::ParseSkipped {
        format: "DOCX".to_string(),
        reason: "element panicked".to_string(),
    };
    let w2 = w.clone();
    assert_eq!(w, w2);
}

#[test]
fn test_convert_result_fields() {
    let result = ConvertResult {
        pdf: vec![0x25, 0x50, 0x44, 0x46],
        warnings: vec![ConvertWarning::UnsupportedElement {
            format: "DOCX".to_string(),
            element: "Image".to_string(),
        }],
        metrics: None,
    };
    assert_eq!(result.pdf, vec![0x25, 0x50, 0x44, 0x46]);
    assert_eq!(result.warnings.len(), 1);
    assert_eq!(result.warnings[0].format(), "DOCX");
}

#[test]
fn test_convert_result_empty_warnings() {
    let result = ConvertResult {
        pdf: vec![1, 2, 3],
        warnings: vec![],
        metrics: None,
    };
    assert!(result.warnings.is_empty());
}

#[test]
fn test_convert_metrics_fields() {
    use std::time::Duration;
    let metrics = ConvertMetrics {
        parse_duration: Duration::from_millis(100),
        codegen_duration: Duration::from_millis(50),
        compile_duration: Duration::from_millis(200),
        total_duration: Duration::from_millis(360),
        input_size_bytes: 1024,
        output_size_bytes: 2048,
        page_count: 5,
    };
    assert_eq!(metrics.parse_duration, Duration::from_millis(100));
    assert_eq!(metrics.codegen_duration, Duration::from_millis(50));
    assert_eq!(metrics.compile_duration, Duration::from_millis(200));
    assert_eq!(metrics.total_duration, Duration::from_millis(360));
    assert_eq!(metrics.input_size_bytes, 1024);
    assert_eq!(metrics.output_size_bytes, 2048);
    assert_eq!(metrics.page_count, 5);
}

#[test]
fn test_convert_metrics_clone() {
    use std::time::Duration;
    let metrics = ConvertMetrics {
        parse_duration: Duration::from_millis(10),
        codegen_duration: Duration::from_millis(20),
        compile_duration: Duration::from_millis(30),
        total_duration: Duration::from_millis(65),
        input_size_bytes: 512,
        output_size_bytes: 1024,
        page_count: 1,
    };
    let cloned = metrics.clone();
    assert_eq!(cloned.parse_duration, metrics.parse_duration);
    assert_eq!(cloned.total_duration, metrics.total_duration);
}

#[test]
fn test_convert_result_with_metrics() {
    use std::time::Duration;
    let result = ConvertResult {
        pdf: vec![0x25, 0x50, 0x44, 0x46],
        warnings: vec![],
        metrics: Some(ConvertMetrics {
            parse_duration: Duration::from_millis(10),
            codegen_duration: Duration::from_millis(20),
            compile_duration: Duration::from_millis(30),
            total_duration: Duration::from_millis(65),
            input_size_bytes: 100,
            output_size_bytes: 200,
            page_count: 1,
        }),
    };
    assert!(result.metrics.is_some());
    let m = result.metrics.unwrap();
    assert_eq!(m.page_count, 1);
}

#[test]
fn test_convert_error_debug_format() {
    let e = ConvertError::UnsupportedFormat("txt".to_string());
    let dbg = format!("{e:?}");
    assert!(dbg.contains("UnsupportedFormat"));
}

#[test]
fn test_unsupported_encryption_display() {
    let e = ConvertError::UnsupportedEncryption;
    let msg = e.to_string();
    assert!(
        msg.contains("encrypted") || msg.contains("password"),
        "UnsupportedEncryption display should mention encryption or password: {msg}"
    );
}

#[test]
fn test_unsupported_encryption_debug() {
    let e = ConvertError::UnsupportedEncryption;
    let dbg = format!("{e:?}");
    assert!(
        dbg.contains("UnsupportedEncryption"),
        "Debug format should contain variant name: {dbg}"
    );
}

#[test]
fn test_all_variants_carry_format() {
    let variants = [
        ConvertWarning::UnsupportedElement {
            format: "DOCX".to_string(),
            element: "x".to_string(),
        },
        ConvertWarning::PartialElement {
            format: "PPTX".to_string(),
            element: "x".to_string(),
            detail: "y".to_string(),
        },
        ConvertWarning::FallbackUsed {
            format: "XLSX".to_string(),
            from: "x".to_string(),
            to: "y".to_string(),
        },
        ConvertWarning::ParseSkipped {
            format: "DOCX".to_string(),
            reason: "x".to_string(),
        },
    ];
    let expected_formats = ["DOCX", "PPTX", "XLSX", "DOCX"];
    for (w, expected) in variants.iter().zip(expected_formats.iter()) {
        assert_eq!(w.format(), *expected);
    }
}
