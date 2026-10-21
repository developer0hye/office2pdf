//! A table width (`CT_TblWidth`) written as a universal measure, a percentage
//! or a fraction reads as the twips or fiftieths of a percent it states, and
//! the properties beside it survive whatever form it takes (issue #2245).

use super::*;
use crate::parser::docx_universal_measures::normalize;
use std::borrow::Cow;
use std::io::{Read, Write};

/// Each `CT_TblWidth` length of [`table_document`], as `w:w` values.
struct TableLengths<'a> {
    table: &'a str,
    first_cell: &'a str,
    second_cell: &'a str,
    cell_margin: &'a str,
    cell_padding: &'a str,
}

/// [`TableLengths`] in twips: the twin every other form is checked against.
const TWIPS: TableLengths<'static> = TableLengths {
    table: "6000",
    first_cell: "1637",
    second_cell: "2674",
    cell_margin: "108",
    cell_padding: "40",
};

/// A two-row table whose first column's cells merge vertically, every cell
/// shaded `5B9BD5` and centred. `grid` adds a `w:tblGrid`; without one the
/// column widths come from the cells' `w:tcW`.
fn table_document(lengths: &TableLengths, width_type: &str, grid: bool) -> String {
    let TableLengths {
        table,
        first_cell,
        second_cell,
        cell_margin,
        cell_padding,
    } = lengths;
    let grid: &str = if grid {
        r#"<w:tblGrid><w:gridCol w:w="1637"/><w:gridCol w:w="2674"/></w:tblGrid>"#
    } else {
        ""
    };
    let cell = |width: &str, merge: &str, text: &str| -> String {
        format!(
            r#"<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="{width_type}"/>{merge}<w:tcMar><w:top w:w="{cell_padding}" w:type="dxa"/></w:tcMar><w:shd w:val="clear" w:color="auto" w:fill="5B9BD5"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>"#
        )
    };
    let rows: String = [
        cell(first_cell, r#"<w:vMerge w:val="restart"/>"#, "Merged")
            + &cell(second_cell, "", "Upper"),
        cell(first_cell, "<w:vMerge/>", "") + &cell(second_cell, "", "Lower"),
    ]
    .iter()
    .map(|cells| format!("<w:tr>{cells}</w:tr>"))
    .collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="{table}" w:type="{width_type}"/><w:tblCellMar><w:left w:w="{cell_margin}" w:type="dxa"/><w:right w:w="{cell_margin}" w:type="dxa"/></w:tblCellMar></w:tblPr>{grid}{rows}</w:tbl><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#
    )
}

fn parse(data: &[u8]) -> Document {
    DocxParser
        .parse(data, &ConvertOptions::default())
        .expect("package parses")
        .0
}

fn typst(data: &[u8]) -> String {
    crate::render::typst_gen::generate_typst(&parse(data))
        .unwrap()
        .source
}

fn assert_reads_as_twips(lengths: &TableLengths) {
    for grid in [true, false] {
        let written = build_docx_with_math(&table_document(lengths, "dxa", grid));
        let twin = build_docx_with_math(&table_document(&TWIPS, "dxa", grid));
        assert_eq!(
            typst(&written),
            typst(&twin),
            "{} (grid: {grid}) should lay out as its twips twin",
            lengths.first_cell
        );
    }
}

#[test]
fn point_table_widths_read_as_their_twips() {
    assert_reads_as_twips(&TableLengths {
        table: "300pt",
        first_cell: "81.85pt",
        second_cell: "133.70pt",
        cell_margin: "5.40pt",
        cell_padding: "2pt",
    });
}

#[test]
fn every_universal_measure_unit_reads_as_its_twips() {
    // 1in = 72pt, 1cm = 72/2.54pt and 1mm a tenth of that, 1pc = 1pi = 12pt.
    // Each value rounds to the twips of TWIPS.
    let cases: [TableLengths; 5] = [
        TableLengths {
            table: "4.1667in",
            first_cell: "1.1368in",
            second_cell: "1.8569in",
            cell_margin: "0.075in",
            cell_padding: "0.0278in",
        },
        TableLengths {
            table: "10.583cm",
            first_cell: "2.8875cm",
            second_cell: "4.7167cm",
            cell_margin: "0.1905cm",
            cell_padding: "0.0706cm",
        },
        TableLengths {
            table: "105.83mm",
            first_cell: "28.875mm",
            second_cell: "47.167mm",
            cell_margin: "1.905mm",
            cell_padding: "0.706mm",
        },
        TableLengths {
            table: "25pc",
            first_cell: "6.8208pc",
            second_cell: "11.1417pc",
            cell_margin: "0.45pc",
            cell_padding: "0.1667pc",
        },
        TableLengths {
            table: "25pi",
            first_cell: "6.8208pi",
            second_cell: "11.1417pi",
            cell_margin: "0.45pi",
            cell_padding: "0.1667pi",
        },
    ];
    for lengths in &cases {
        assert_reads_as_twips(lengths);
    }
}

#[test]
fn fractional_twips_round_to_the_nearest_twip() {
    assert_reads_as_twips(&TableLengths {
        table: "5999.6",
        first_cell: "1636.5",
        second_cell: "2673.9",
        cell_margin: "108.4",
        cell_padding: "39.5",
    });
}

/// Without a `w:tblGrid` the columns are the cells' own widths, so the
/// value each unit converts to shows directly.
#[test]
fn universal_measure_cell_widths_size_the_columns() {
    let cases: [(&str, &str, [f64; 2]); 3] = [
        ("81.85pt", "133.70pt", [81.85, 133.7]),
        ("1.25in", "2in", [90.0, 144.0]),
        // 3.2cm is 1814.17 twips and 5cm 2834.65, read as 1814 and 2835.
        ("3.2cm", "5cm", [90.7, 141.75]),
    ];
    for (first_cell, second_cell, expected) in cases {
        let lengths = TableLengths {
            first_cell,
            second_cell,
            ..TWIPS
        };
        let doc = parse(&build_docx_with_math(&table_document(
            &lengths, "dxa", false,
        )));
        let widths: &[f64] = &first_table(&doc).column_widths;
        assert_eq!(widths.len(), 2, "{first_cell}: {widths:?}");
        for (width, expected) in widths.iter().zip(expected) {
            assert!(
                (width - expected).abs() < 1e-6,
                "{first_cell}: column widths {widths:?}, expected {expected}"
            );
        }
    }
}

/// The cell's shading, vertical alignment and merge are its `w:tcPr`'s other
/// children, which the whole element used to be dropped with.
#[test]
fn a_cell_keeps_its_properties_whatever_form_its_width_takes() {
    let shading = Color::new(0x5B, 0x9B, 0xD5);
    for (first_cell, width_type) in [
        ("81.85pt", "dxa"),
        ("1.25in", "dxa"),
        ("1637.4", "dxa"),
        ("50%", "pct"),
    ] {
        let lengths = TableLengths {
            first_cell,
            ..TWIPS
        };
        let doc = parse(&build_docx_with_math(&table_document(
            &lengths, width_type, true,
        )));
        let merged: &TableCell = &first_table(&doc).rows[0].cells[0];
        assert_eq!(merged.background, Some(shading), "{first_cell}");
        assert_eq!(
            merged.vertical_align,
            Some(CellVerticalAlign::Center),
            "{first_cell}"
        );
        assert_eq!(merged.row_span, 2, "{first_cell}");
    }
}

/// A percentage reads as the fiftieths of a percent a Transitional package
/// writes for it.
#[test]
fn percentage_table_widths_read_as_fiftieths_of_a_percent() {
    let percent = TableLengths {
        table: "100%",
        first_cell: "40%",
        second_cell: "60%",
        ..TWIPS
    };
    let fiftieths = TableLengths {
        table: "5000",
        first_cell: "2000",
        second_cell: "3000",
        ..TWIPS
    };
    for grid in [true, false] {
        assert_eq!(
            typst(&build_docx_with_math(&table_document(
                &percent, "pct", grid
            ))),
            typst(&build_docx_with_math(&table_document(
                &fiftieths, "pct", grid
            ))),
            "grid: {grid}"
        );
    }
}

/// A table style's widths sit in the styles part.
#[test]
fn a_table_style_reads_its_universal_measure_cell_margins() {
    let styles = |margin: &str| -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="table" w:styleId="Grid"><w:name w:val="Grid"/><w:tblPr><w:tblCellMar><w:top w:w="{margin}" w:type="dxa"/><w:left w:w="{margin}" w:type="dxa"/><w:bottom w:w="{margin}" w:type="dxa"/><w:right w:w="{margin}" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>"#
        )
    };
    let document: String = table_document(&TWIPS, "dxa", true)
        .replace("<w:tblPr>", r#"<w:tblPr><w:tblStyle w:val="Grid"/>"#)
        .replace(
            r#"<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>"#,
            "",
        );
    for (written, twips) in [("14.4pt", "288"), ("0.5cm", "283")] {
        let doc = parse(&build_docx_with_styles_xml(&document, &styles(written)));
        let twin = parse(&build_docx_with_styles_xml(&document, &styles(twips)));
        let padding = first_table(&doc).default_cell_padding;
        assert_eq!(
            padding,
            first_table(&twin).default_cell_padding,
            "{written}"
        );
        assert!(
            padding.is_some_and(|inset| inset.left > 10.0),
            "{written}: {padding:?}"
        );
    }
}

fn zip_parts(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, content) in parts {
        let options = zip::write::SimpleFileOptions::default();
        zip.start_file(*name, options).unwrap();
        zip.write_all(content).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn read_part(package: &[u8], name: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(package)).unwrap();
    let mut part = Vec::new();
    archive
        .by_name(name)
        .unwrap()
        .read_to_end(&mut part)
        .unwrap();
    part
}

/// Tables also live in headers, footers and notes; the rewrite reaches every
/// part, changes only the widths, and copies the rest byte for byte.
#[test]
fn widths_are_rewritten_in_every_part_and_nothing_else_changes() {
    let header: &[u8] = br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tbl><w:tblPr><w:tblW w:w="3in" w:type="dxa"/><w:tblInd w:w="-5.4pt" w:type="dxa"/></w:tblPr><w:tr><w:trPr><w:wBefore w:w="1cm" w:type="dxa"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="81.85pt" w:type="dxa"/><w:tcMar><w:start w:w="2mm" w:type="dxa"/></w:tcMar><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders></w:tcPr><w:p><w:r><w:t xml:space="preserve">72pt &amp; 50% </w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:hdr>"#;
    let expected: &[u8] = br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblInd w:w="-108" w:type="dxa"/></w:tblPr><w:tr><w:trPr><w:wBefore w:w="567" w:type="dxa"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="1637" w:type="dxa"/><w:tcMar><w:start w:w="113" w:type="dxa"/></w:tcMar><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders></w:tcPr><w:p><w:r><w:t xml:space="preserve">72pt &amp; 50% </w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:hdr>"#;
    let document: String = table_document(
        &TableLengths {
            first_cell: "81.85pt",
            ..TWIPS
        },
        "dxa",
        true,
    );
    let image: &[u8] = b"\x89PNG not really w:w=\"1pt\"";
    let package: Vec<u8> = zip_parts(&[
        ("word/document.xml", document.as_bytes()),
        ("word/header1.xml", header),
        ("word/media/image1.png", image),
    ]);

    let normalized: Cow<'_, [u8]> = normalize(&package);
    assert!(matches!(normalized, Cow::Owned(_)));
    assert_eq!(
        String::from_utf8(read_part(&normalized, "word/header1.xml")).unwrap(),
        String::from_utf8(expected.to_vec()).unwrap()
    );
    assert_eq!(
        String::from_utf8(read_part(&normalized, "word/document.xml")).unwrap(),
        document.replace(r#"w:w="81.85pt""#, r#"w:w="1637""#)
    );
    assert_eq!(read_part(&normalized, "word/media/image1.png"), image);
}

#[test]
fn a_package_with_twips_widths_passes_through_unchanged() {
    let package: Vec<u8> = build_docx_with_math(&table_document(&TWIPS, "dxa", true));
    assert!(matches!(normalize(&package), Cow::Borrowed(_)));
}

fn cell_backgrounds(blocks: &[Block], out: &mut Vec<(Color, u32)>) {
    for block in blocks {
        if let Block::Table(table) = block {
            for cell in table.rows.iter().flat_map(|row| &row.cells) {
                if let Some(background) = cell.background {
                    out.push((background, cell.row_span));
                }
                cell_backgrounds(&cell.content, out);
            }
        }
    }
}

/// The tracked Strict fixture the issue was found on: every `w:tcW` is in
/// points, and its cells' fills and the nested table's merge were lost.
#[test]
fn strict_fixture_paints_its_cell_fills() {
    let data: &[u8] = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../tests/fixtures/docx/libreoffice/tdf79272_strictDxa.docx"
    ));
    let doc = parse(data);
    let mut backgrounds: Vec<(Color, u32)> = Vec::new();
    for page in &doc.pages {
        if let Page::Flow(page) = page {
            cell_backgrounds(&page.content, &mut backgrounds);
        }
    }
    let count = |r: u8, g: u8, b: u8, row_span: u32| -> usize {
        backgrounds
            .iter()
            .filter(|&&(color, span)| color == Color::new(r, g, b) && span == row_span)
            .count()
    };
    assert_eq!(count(0x5B, 0x9B, 0xD5, 1), 2, "{backgrounds:?}");
    assert_eq!(count(0x9C, 0xC2, 0xE5, 3), 1, "{backgrounds:?}");
    assert_eq!(count(0xBD, 0xD6, 0xEE, 1), 3, "{backgrounds:?}");
    assert_eq!(count(0xFF, 0xC0, 0x00, 1), 3, "{backgrounds:?}");
}
