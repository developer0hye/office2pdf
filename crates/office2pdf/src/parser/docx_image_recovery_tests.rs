use super::*;
use std::io::{Read, Write};

const VALID_REPORT: &[u8] =
    include_bytes!("../../../../tests/fixtures/docx/docxide/valid_png_report.docx");
const TRUNCATED_REPORT: &[u8] =
    include_bytes!("../../../../tests/fixtures/docx/docxide/truncated_png_report.docx");

fn report_with_png(png: &[u8]) -> Vec<u8> {
    replace_package_png(VALID_REPORT, png)
}

fn replace_package_png(package: &[u8], png: &[u8]) -> Vec<u8> {
    let mut source = zip::ZipArchive::new(Cursor::new(package)).unwrap();
    let mut output = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for index in 0..source.len() {
        let mut entry = source.by_index(index).unwrap();
        let mut bytes: Vec<u8> = Vec::new();
        entry.read_to_end(&mut bytes).unwrap();
        let name = entry.name().expect("entry name should decode").into_owned();
        output
            .start_file(&name, zip::write::SimpleFileOptions::default())
            .unwrap();
        output
            .write_all(
                if name.starts_with("word/media/") && name.ends_with(".png") {
                    png
                } else {
                    &bytes
                },
            )
            .unwrap();
    }
    output.finish().unwrap().into_inner()
}

#[test]
fn docx_damaged_png_header_and_footer_pictures_keep_their_extents() {
    let read_png = |report| {
        let mut archive = zip::ZipArchive::new(Cursor::new(report)).unwrap();
        let mut data: Vec<u8> = Vec::new();
        archive
            .by_name("word/media/image1.png")
            .unwrap()
            .read_to_end(&mut data)
            .unwrap();
        data
    };
    let valid = read_png(VALID_REPORT);
    let damaged = read_png(TRUNCATED_REPORT);
    let picture_paragraph = || {
        docx_rs::Paragraph::new().add_run(
            docx_rs::Run::new().add_image(docx_rs::Pic::new(&valid).size(1_524_000, 762_000)),
        )
    };
    let docx = docx_rs::Docx::new()
        .header(docx_rs::Header::new().add_paragraph(picture_paragraph()))
        .footer(docx_rs::Footer::new().add_paragraph(picture_paragraph()))
        .add_paragraph(
            docx_rs::Paragraph::new().add_run(docx_rs::Run::new().add_text("Report body")),
        );
    let mut cursor = Cursor::new(Vec::new());
    docx.build().pack(&mut cursor).unwrap();
    let bytes = replace_package_png(&cursor.into_inner(), &damaged);
    let (document, warnings) = DocxParser
        .parse(&bytes, &ConvertOptions::default())
        .unwrap();
    let Page::Flow(page) = &document.pages[0] else {
        panic!("flow section")
    };
    for story in [page.header.as_ref().unwrap(), page.footer.as_ref().unwrap()] {
        let image = story
            .paragraphs
            .iter()
            .flat_map(|paragraph| &paragraph.elements)
            .find_map(|element| match element {
                HFInline::Image(image) => Some(image),
                _ => None,
            })
            .expect("damaged story picture must keep its layout space");
        assert_eq!((image.width, image.height), (Some(120.0), Some(60.0)));
        assert!(image::load_from_memory(&image.data).is_ok());
    }
    assert!(!warnings.is_empty());
}

fn report_picture(document: &Document) -> &ImageData {
    all_blocks(document)
        .iter()
        .find_map(|block| match block {
            Block::Image(image) => Some(image),
            _ => None,
        })
        .expect("picture layout space retained")
}

#[test]
fn docx_damaged_png_crc_recovery_preserves_color_alpha_and_grayscale_samples() {
    for (color, channels) in [
        (png::ColorType::Rgb, 3),
        (png::ColorType::Rgba, 4),
        (png::ColorType::Grayscale, 1),
    ] {
        let samples: Vec<u8> = (0..24 * 12 * channels)
            .map(|index| ((index * 31 + 17) % 256) as u8)
            .collect();
        let mut original: Vec<u8> = Vec::new();
        {
            let mut encoder = png::Encoder::new(&mut original, 24, 12);
            encoder.set_color(color);
            encoder.set_depth(png::BitDepth::Eight);
            encoder
                .write_header()
                .unwrap()
                .write_image_data(&samples)
                .unwrap();
        }
        let expected = image::load_from_memory(&original).unwrap().to_rgba8();
        let mut damaged: Vec<u8> = original.clone();
        let mut offset: usize = 8;
        while offset + 12 <= damaged.len() {
            let size = u32::from_be_bytes(damaged[offset..offset + 4].try_into().unwrap()) as usize;
            if &damaged[offset + 4..offset + 8] == b"IDAT" {
                damaged[offset + 8 + size] ^= 1;
                break;
            }
            offset += size + 12;
        }
        assert!(image::load_from_memory(&damaged).is_err());
        let (document, warnings) = DocxParser
            .parse(&report_with_png(&damaged), &ConvertOptions::default())
            .unwrap();
        let actual = image::load_from_memory(&report_picture(&document).data)
            .unwrap()
            .to_rgba8();
        assert_eq!(actual, expected, "{color:?}");
        assert!(!warnings.is_empty());

        let (valid_document, warnings) = DocxParser
            .parse(&report_with_png(&original), &ConvertOptions::default())
            .unwrap();
        assert_eq!(report_picture(&valid_document).data, original);
        assert!(warnings.is_empty());
    }
}

#[test]
fn docx_damaged_png_retains_a_decodable_picture_and_declared_extent() {
    let (document, warnings) = DocxParser
        .parse(TRUNCATED_REPORT, &ConvertOptions::default())
        .unwrap();
    let images: Vec<&ImageData> = all_blocks(&document)
        .iter()
        .filter_map(|block| match block {
            Block::Image(image) => Some(image),
            _ => None,
        })
        .collect();
    assert_eq!(images.len(), 1, "the picture must keep its layout space");
    assert_eq!(images[0].width, Some(120.0));
    assert_eq!(images[0].height, Some(60.0));
    assert!(
        image::load_from_memory(&images[0].data).is_ok(),
        "damaged media must not reach the PDF decoder"
    );
    assert!(warnings.iter().any(|warning| matches!(
        warning,
        ConvertWarning::PartialElement { format, .. }
            | ConvertWarning::FallbackUsed { format, .. } if format == "DOCX"
    )));
}

#[test]
fn docx_damaged_png_truncated_tail_preserves_complete_decoded_rows() {
    for (color, channels) in [
        (png::ColorType::Rgb, 3),
        (png::ColorType::Rgba, 4),
        (png::ColorType::Grayscale, 1),
    ] {
        let samples: Vec<u8> = (0..24 * 12 * channels)
            .map(|index| ((index * 31 + 17) % 255 + 1) as u8)
            .collect();
        let mut original: Vec<u8> = Vec::new();
        {
            let mut encoder = png::Encoder::new(&mut original, 24, 12);
            encoder.set_color(color);
            encoder.set_depth(png::BitDepth::Eight);
            encoder.set_compression(png::Compression::NoCompression);
            encoder.set_filter(png::Filter::NoFilter);
            encoder
                .write_header()
                .unwrap()
                .write_image_data(&samples)
                .unwrap();
        }
        let expected = image::load_from_memory(&original).unwrap().to_rgba8();
        let mut offset: usize = 8;
        while &original[offset + 4..offset + 8] != b"IDAT" {
            let size = u32::from_be_bytes(original[offset..offset + 4].try_into().unwrap());
            offset += size as usize + 12;
        }
        let size = u32::from_be_bytes(original[offset..offset + 4].try_into().unwrap()) as usize;
        let damaged = &original[..offset + 8 + size / 2];
        assert!(image::load_from_memory(damaged).is_err());
        let (document, warnings) = DocxParser
            .parse(&report_with_png(damaged), &ConvertOptions::default())
            .unwrap();
        let actual = image::load_from_memory(&report_picture(&document).data)
            .unwrap()
            .to_rgba8();
        assert_eq!(actual.dimensions(), expected.dimensions());
        for y in 0..3 {
            for x in 0..24 {
                assert_eq!(
                    actual.get_pixel(x, y),
                    expected.get_pixel(x, y),
                    "{color:?}"
                );
            }
        }
        let missing = if channels == 4 {
            image::Rgba([0, 0, 0, 0])
        } else {
            image::Rgba([0, 0, 0, 255])
        };
        assert_eq!(*actual.get_pixel(23, 11), missing, "{color:?}");
        assert!(!warnings.is_empty());
    }
}

#[cfg(not(target_arch = "wasm32"))]
#[test]
fn docx_damaged_png_converts_the_report_and_preserves_surrounding_text() {
    for data in [VALID_REPORT, TRUNCATED_REPORT] {
        let result = crate::convert_bytes(data, crate::Format::Docx, &ConvertOptions::default())
            .expect("one damaged picture must not abort a report");
        let text: String = pdf_extract::extract_text_from_mem(&result.pdf).unwrap();
        assert!(text.contains("Quarterly chart:"));
        assert!(text.contains("See the appendix for the complete report."));
    }
}
