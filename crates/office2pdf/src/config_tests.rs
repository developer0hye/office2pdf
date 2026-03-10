use super::*;

#[test]
fn test_format_from_extension() {
    assert_eq!(Format::from_extension("docx"), Some(Format::Docx));
    assert_eq!(Format::from_extension("DOCX"), Some(Format::Docx));
    assert_eq!(Format::from_extension("pptx"), Some(Format::Pptx));
    assert_eq!(Format::from_extension("xlsx"), Some(Format::Xlsx));
    assert_eq!(Format::from_extension("pdf"), None);
    assert_eq!(Format::from_extension("txt"), None);
}

#[test]
fn test_slide_range_single() {
    let r = SlideRange::parse("3").unwrap();
    assert_eq!(r.start, 3);
    assert_eq!(r.end, 3);
    assert!(!r.contains(2));
    assert!(r.contains(3));
    assert!(!r.contains(4));
}

#[test]
fn test_slide_range_range() {
    let r = SlideRange::parse("2-5").unwrap();
    assert_eq!(r.start, 2);
    assert_eq!(r.end, 5);
    assert!(!r.contains(1));
    assert!(r.contains(2));
    assert!(r.contains(3));
    assert!(r.contains(5));
    assert!(!r.contains(6));
}

#[test]
fn test_slide_range_parse_errors() {
    assert!(SlideRange::parse("abc").is_err());
    assert!(SlideRange::parse("0").is_err());
    assert!(SlideRange::parse("5-2").is_err());
    assert!(SlideRange::parse("0-3").is_err());
    assert!(SlideRange::parse("a-b").is_err());
}

#[test]
fn test_convert_options_default() {
    let opts = ConvertOptions::default();
    assert!(opts.sheet_names.is_none());
    assert!(opts.slide_range.is_none());
}

#[test]
fn test_convert_options_with_sheets() {
    let opts = ConvertOptions {
        sheet_names: Some(vec!["Sheet1".to_string(), "Data".to_string()]),
        ..Default::default()
    };
    assert_eq!(opts.sheet_names.as_ref().unwrap().len(), 2);
}

#[test]
fn test_convert_options_with_slide_range() {
    let opts = ConvertOptions {
        slide_range: Some(SlideRange::new(1, 3)),
        ..Default::default()
    };
    assert!(opts.slide_range.as_ref().unwrap().contains(2));
}

#[test]
fn test_pdf_standard_enum_exists() {
    let std = PdfStandard::PdfA2b;
    assert_eq!(format!("{std:?}"), "PdfA2b");
}

#[test]
fn test_convert_options_pdf_standard_default_none() {
    let opts = ConvertOptions::default();
    assert!(opts.pdf_standard.is_none());
}

#[test]
fn test_convert_options_with_pdf_standard() {
    let opts = ConvertOptions {
        pdf_standard: Some(PdfStandard::PdfA2b),
        ..Default::default()
    };
    assert_eq!(opts.pdf_standard, Some(PdfStandard::PdfA2b));
}

// --- PaperSize tests ---

#[test]
fn test_paper_size_a4_dimensions() {
    let (w, h) = PaperSize::A4.dimensions();
    assert!((w - 595.28).abs() < 0.01);
    assert!((h - 841.89).abs() < 0.01);
}

#[test]
fn test_paper_size_letter_dimensions() {
    let (w, h) = PaperSize::Letter.dimensions();
    assert!((w - 612.0).abs() < 0.01);
    assert!((h - 792.0).abs() < 0.01);
}

#[test]
fn test_paper_size_legal_dimensions() {
    let (w, h) = PaperSize::Legal.dimensions();
    assert!((w - 612.0).abs() < 0.01);
    assert!((h - 1008.0).abs() < 0.01);
}

#[test]
fn test_paper_size_custom_dimensions() {
    let ps = PaperSize::Custom {
        width: 400.0,
        height: 600.0,
    };
    assert_eq!(ps.dimensions(), (400.0, 600.0));
}

#[test]
fn test_paper_size_parse() {
    assert_eq!(PaperSize::parse("a4").unwrap(), PaperSize::A4);
    assert_eq!(PaperSize::parse("A4").unwrap(), PaperSize::A4);
    assert_eq!(PaperSize::parse("letter").unwrap(), PaperSize::Letter);
    assert_eq!(PaperSize::parse("LETTER").unwrap(), PaperSize::Letter);
    assert_eq!(PaperSize::parse("legal").unwrap(), PaperSize::Legal);
    assert!(PaperSize::parse("tabloid").is_err());
}

#[test]
fn test_convert_options_paper_size_default_none() {
    let opts = ConvertOptions::default();
    assert!(opts.paper_size.is_none());
}

#[test]
fn test_convert_options_font_paths_default_empty() {
    let opts = ConvertOptions::default();
    assert!(opts.font_paths.is_empty());
}

#[test]
fn test_convert_options_landscape_default_none() {
    let opts = ConvertOptions::default();
    assert!(opts.landscape.is_none());
}

#[test]
fn test_convert_options_with_paper_size() {
    let opts = ConvertOptions {
        paper_size: Some(PaperSize::Letter),
        ..Default::default()
    };
    assert_eq!(opts.paper_size, Some(PaperSize::Letter));
}

#[test]
fn test_convert_options_with_font_paths() {
    let opts = ConvertOptions {
        font_paths: vec![
            std::path::PathBuf::from("/usr/share/fonts"),
            std::path::PathBuf::from("/home/user/.fonts"),
        ],
        ..Default::default()
    };
    assert_eq!(opts.font_paths.len(), 2);
}

#[test]
fn test_convert_options_with_landscape() {
    let opts = ConvertOptions {
        landscape: Some(true),
        ..Default::default()
    };
    assert_eq!(opts.landscape, Some(true));
}

#[test]
fn test_convert_options_tagged_default_false() {
    let opts = ConvertOptions::default();
    assert!(!opts.tagged);
}

#[test]
fn test_convert_options_pdf_ua_default_false() {
    let opts = ConvertOptions::default();
    assert!(!opts.pdf_ua);
}

#[test]
fn test_convert_options_with_tagged() {
    let opts = ConvertOptions {
        tagged: true,
        ..Default::default()
    };
    assert!(opts.tagged);
}

#[test]
fn test_convert_options_with_pdf_ua() {
    let opts = ConvertOptions {
        pdf_ua: true,
        ..Default::default()
    };
    assert!(opts.pdf_ua);
}

#[test]
fn test_convert_options_streaming_default_false() {
    let opts = ConvertOptions::default();
    assert!(!opts.streaming);
}

#[test]
fn test_convert_options_streaming_chunk_size_default_none() {
    let opts = ConvertOptions::default();
    assert!(opts.streaming_chunk_size.is_none());
}

#[test]
fn test_convert_options_with_streaming() {
    let opts = ConvertOptions {
        streaming: true,
        ..Default::default()
    };
    assert!(opts.streaming);
}

#[test]
fn test_convert_options_with_streaming_chunk_size() {
    let opts = ConvertOptions {
        streaming: true,
        streaming_chunk_size: Some(500),
        ..Default::default()
    };
    assert!(opts.streaming);
    assert_eq!(opts.streaming_chunk_size, Some(500));
}
