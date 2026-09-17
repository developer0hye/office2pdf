//! A cell alignment's `indent` level (issue #1109).
//!
//! Every expected point count here comes from the native Excel for Mac probe
//! series recorded on that issue: the unit is three spaces of the workbook
//! Normal font, each rounded to a whole point.

use super::*;

/// Build a workbook whose `xl/styles.xml` and worksheet body are both under
/// the test's control. umya's builder cannot express an `indent`, and the
/// `cellXfs` index a cell's `s` attribute names is exactly what this joins.
fn build_xlsx_with_styles_and_sheet(styles_xml: &str, sheet_body: &str) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::FileOptions::default();
    let mut write = |path: &str, body: &str| {
        zip.start_file(path, options).unwrap();
        std::io::Write::write_all(&mut zip, body.as_bytes()).unwrap();
    };

    write(
        "[Content_Types].xml",
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#,
    );
    write(
        "_rels/.rels",
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#,
    );
    write(
        "xl/workbook.xml",
        r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#,
    );
    write(
        "xl/_rels/workbook.xml.rels",
        r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    );
    write("xl/styles.xml", styles_xml);
    write(
        "xl/worksheets/sheet1.xml",
        &format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
{sheet_body}
</worksheet>"#
        ),
    );

    zip.finish().unwrap().into_inner()
}

/// A stylesheet whose Normal font is `normal_family` at `normal_size_pt`, with
/// a second font fixed at Calibri 11 for the cells, and one `cellXfs` entry
/// per fragment in `alignments` (entry 0 stays the unstyled default).
fn styles_with_alignments(normal_family: &str, normal_size_pt: f64, alignments: &[&str]) -> String {
    let cell_xfs: String = alignments
        .iter()
        .map(|alignment| {
            format!(
                r#"<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">{alignment}</xf>"#
            )
        })
        .collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font><sz val="{normal_size_pt}"/><name val="{normal_family}"/></font>
<font><sz val="11"/><name val="Calibri"/></font>
</fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="{count}">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
{cell_xfs}
</cellXfs>
</styleSheet>"#,
        count = alignments.len() + 1,
    )
}

/// One row of inline-string cells, each naming the `cellXfs` entry beside it.
fn sheet_with_cells(cells: &[(&str, u32, &str)]) -> String {
    let cell_xml: String = cells
        .iter()
        .map(|(reference, style, text)| {
            format!(r#"<c r="{reference}" s="{style}" t="inlineStr"><is><t>{text}</t></is></c>"#)
        })
        .collect();
    format!("<sheetData><row r=\"1\">{cell_xml}</row></sheetData>")
}

/// The padding of the first cell of the first row of the parsed workbook.
fn first_cell_padding(data: &[u8]) -> Option<crate::ir::Insets> {
    let (doc, _warnings) = XlsxParser
        .parse(data, &ConvertOptions::default())
        .expect("workbook should parse");
    let tp = get_sheet_page(&doc, 0);
    tp.table.rows[0].cells[0].padding
}

/// Three spaces of Calibri 11 (0.2261em -> 2.49pt, a 2pt whole-point advance)
/// is the 6pt/level the probe measures.
const CALIBRI_11_INDENT_UNIT_PT: f64 = 6.0;

#[test]
fn test_left_indent_insets_the_text_by_three_normal_font_spaces_per_level() {
    let data = build_xlsx_with_styles_and_sheet(
        &styles_with_alignments(
            "Calibri",
            11.0,
            &[r#"<alignment horizontal="left" indent="2"/>"#],
        ),
        &sheet_with_cells(&[("A1", 1, "Indented")]),
    );

    let padding = first_cell_padding(&data).expect("an indented cell states its own padding");
    assert!(
        (padding.left - (XLSX_CELL_PADDING.left + 2.0 * CALIBRI_11_INDENT_UNIT_PT)).abs() < 0.01,
        "two levels should inset the text 12pt past the 3pt cell inset, got {}",
        padding.left
    );
    assert!(
        (padding.right - XLSX_CELL_PADDING.right).abs() < 0.01,
        "a left indent leaves the right inset alone, got {}",
        padding.right
    );
}

#[test]
fn test_right_aligned_indent_insets_from_the_right_edge() {
    let data = build_xlsx_with_styles_and_sheet(
        &styles_with_alignments(
            "Calibri",
            11.0,
            &[r#"<alignment horizontal="right" indent="1"/>"#],
        ),
        &sheet_with_cells(&[("A1", 1, "Indented")]),
    );

    let padding = first_cell_padding(&data).expect("an indented cell states its own padding");
    assert!(
        (padding.right - (XLSX_CELL_PADDING.right + CALIBRI_11_INDENT_UNIT_PT)).abs() < 0.01,
        "a right-aligned indent moves the text in from the right, got {}",
        padding.right
    );
    assert!(
        (padding.left - XLSX_CELL_PADDING.left).abs() < 0.01,
        "the left inset stays at the cell inset, got {}",
        padding.left
    );
}

#[test]
fn test_general_aligned_number_indents_from_the_right_edge() {
    let styles = styles_with_alignments("Calibri", 11.0, &[r#"<alignment indent="1"/>"#]);
    let data = build_xlsx_with_styles_and_sheet(
        &styles,
        r#"<sheetData><row r="1"><c r="A1" s="1"><v>42</v></c></row></sheetData>"#,
    );

    let padding = first_cell_padding(&data).expect("an indented cell states its own padding");
    assert!(
        (padding.right - (XLSX_CELL_PADDING.right + CALIBRI_11_INDENT_UNIT_PT)).abs() < 0.01,
        "a general-aligned number right-aligns, so its indent comes off the right, got {}",
        padding.right
    );
}

#[test]
fn test_indent_unit_ignores_the_cell_font_and_follows_the_workbook_normal_font() {
    // The cell font is Calibri 22, four times the Normal font's own space, and
    // the probe series moved 8pt, 11pt and 22pt cells by the same 6pt/level.
    let styles = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font><sz val="11"/><name val="Calibri"/></font>
<font><sz val="22"/><name val="Calibri"/></font>
</fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" indent="1"/></xf>
</cellXfs>
</styleSheet>"#;
    let data =
        build_xlsx_with_styles_and_sheet(styles, &sheet_with_cells(&[("A1", 1, "Indented")]));

    // The cell's own text box is what the indent stacks on, and that box does
    // follow the cell font (issue #1165); the level's own unit does not.
    let cell_box_left_pt: f64 =
        super::super::xlsx_cells::cell_left_inset_pt("Calibri", 22.0, false);
    let padding = first_cell_padding(&data).expect("an indented cell states its own padding");
    assert!(
        (padding.left - (cell_box_left_pt + CALIBRI_11_INDENT_UNIT_PT)).abs() < 0.01,
        "a 22pt cell under an 11pt Normal font still indents by the Normal font's unit, got {}",
        padding.left
    );
}

#[test]
fn test_indent_unit_scales_with_the_normal_font_size() {
    // Calibri 16: 3.62pt of space rounds to 4pt, so three of them are 12pt.
    let data = build_xlsx_with_styles_and_sheet(
        &styles_with_alignments(
            "Calibri",
            16.0,
            &[r#"<alignment horizontal="left" indent="1"/>"#],
        ),
        &sheet_with_cells(&[("A1", 1, "Indented")]),
    );

    let padding = first_cell_padding(&data).expect("an indented cell states its own padding");
    assert!(
        (padding.left - (XLSX_CELL_PADDING.left + 12.0)).abs() < 0.01,
        "Calibri 16 measures a 12pt indent unit, got {}",
        padding.left
    );
}

#[test]
fn test_indent_unit_follows_the_normal_font_face() {
    // Courier New 11: 6.60pt of space rounds to 7pt, so three of them are 21pt.
    let data = build_xlsx_with_styles_and_sheet(
        &styles_with_alignments(
            "Courier New",
            11.0,
            &[r#"<alignment horizontal="left" indent="1"/>"#],
        ),
        &sheet_with_cells(&[("A1", 1, "Indented")]),
    );

    let padding = first_cell_padding(&data).expect("an indented cell states its own padding");
    assert!(
        (padding.left - (XLSX_CELL_PADDING.left + 21.0)).abs() < 0.01,
        "Courier New 11 measures a 21pt indent unit, got {}",
        padding.left
    );
}

#[test]
fn segoe_ui_normal_font_uses_its_native_three_space_indent() {
    // Gift Budget and Tracker's Normal style is Segoe UI 10 and its wrapped
    // B4:D4 instruction panel declares indent="2". Native Excel gives one
    // level 9pt: round(0.273926em * 10pt) * 3 spaces. Falling back to
    // Calibri's 6pt unit widens the panel by 6pt and changes every wrap after
    // the third line (issue #1472).
    let data = build_xlsx_with_styles_and_sheet(
        &styles_with_alignments(
            "Segoe UI",
            10.0,
            &[r#"<alignment horizontal="left" indent="2"/>"#],
        ),
        &sheet_with_cells(&[("A1", 1, "Indented")]),
    );

    let padding = first_cell_padding(&data).expect("an indented cell states its own padding");
    assert!(
        (padding.left - (XLSX_CELL_PADDING.left + 18.0)).abs() < 0.01,
        "two native Segoe UI indent levels should add 18pt, got {}",
        padding.left
    );
}

#[test]
fn test_indent_applies_even_when_the_xf_switches_alignment_off() {
    // `applyAlignment="false"` beside an indent is what LibreOffice writes,
    // and the reported workbook's default `cellXfs[0]` carries exactly that.
    // The native export indents such a cell all the same.
    let styles = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="false"><alignment horizontal="left" indent="2"/></xf>
</cellXfs>
</styleSheet>"#;
    let data =
        build_xlsx_with_styles_and_sheet(styles, &sheet_with_cells(&[("A1", 1, "Indented")]));

    let padding = first_cell_padding(&data).expect("an indented cell states its own padding");
    assert!(
        (padding.left - (XLSX_CELL_PADDING.left + 2.0 * CALIBRI_11_INDENT_UNIT_PT)).abs() < 0.01,
        "applyAlignment=false does not suppress the indent, got {}",
        padding.left
    );
}

#[test]
fn test_indent_is_not_inherited_from_the_cell_style_xf() {
    // A `cellXfs` entry with no alignment of its own, pointing at a
    // `cellStyleXfs` entry that indents by 2: the native export leaves the
    // cell flush, so the level is the cell xf's own or nothing.
    let styles = r#"<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="left" indent="2"/></xf>
</cellStyleXfs>
<cellXfs count="2">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1"/>
</cellXfs>
</styleSheet>"#;
    let data = build_xlsx_with_styles_and_sheet(styles, &sheet_with_cells(&[("A1", 1, "Flush")]));

    let padding = first_cell_padding(&data);
    assert!(
        padding.is_none_or(|insets| (insets.left - XLSX_CELL_PADDING.left).abs() < 0.01),
        "a style xf's indent does not reach the cell, got {padding:?}"
    );
}

#[test]
fn test_an_unindented_cell_keeps_the_sheet_default_padding() {
    let data = build_xlsx_with_styles_and_sheet(
        &styles_with_alignments(
            "Calibri",
            11.0,
            &[r#"<alignment horizontal="left" indent="0"/>"#],
        ),
        &sheet_with_cells(&[("A1", 1, "Flush")]),
    );

    let padding = first_cell_padding(&data);
    assert!(
        padding.is_none_or(|insets| (insets.left - XLSX_CELL_PADDING.left).abs() < 0.01),
        "indent=0 leaves the cell on the sheet default, got {padding:?}"
    );
}

#[test]
fn test_a_cell_without_its_own_style_takes_the_row_format_indent() {
    // Excel resolves a `<c>` with no `s` through its row's format, then its
    // column's; a row-formatted band of indented cells is written that way.
    let styles = styles_with_alignments(
        "Calibri",
        11.0,
        &[r#"<alignment horizontal="left" indent="1"/>"#],
    );
    let data = build_xlsx_with_styles_and_sheet(
        &styles,
        r#"<sheetData><row r="1" s="1" customFormat="1"><c r="A1" t="inlineStr"><is><t>Indented</t></is></c></row></sheetData>"#,
    );

    let padding = first_cell_padding(&data).expect("an indented cell states its own padding");
    assert!(
        (padding.left - (XLSX_CELL_PADDING.left + CALIBRI_11_INDENT_UNIT_PT)).abs() < 0.01,
        "the row's own format supplies the indent, got {}",
        padding.left
    );
}

#[test]
fn test_indent_narrows_the_width_a_line_has_before_it_spills() {
    // The indent takes its points out of the cell's own text width, so a line
    // that fits flush no longer fits indented and Excel paints it on into the
    // empty neighbour. Measured on the probe series as a wrap: a string that
    // stayed on one line flush broke in two at indent 2.
    //
    // The string is chosen to clear both thresholds by several points, so the
    // test tracks the behaviour rather than a rounding edge: `Nine chars`
    // prices at 47.26pt against the 55pt this 60pt column leaves flush and the
    // 43pt it leaves at indent 2. The earlier `Nine char` (42.65pt) cleared the
    // indented threshold by 0.35pt and stopped spilling the moment the right
    // inset narrowed by a point (issue #1157).
    let styles = styles_with_alignments(
        "Calibri",
        11.0,
        &[
            r#"<alignment horizontal="left" indent="0"/>"#,
            r#"<alignment horizontal="left" indent="2"/>"#,
        ],
    );
    let columns = r#"<cols><col min="1" max="2" width="10" customWidth="1"/></cols>"#;
    let flush = build_xlsx_with_styles_and_sheet(
        &styles,
        &format!(
            "{columns}<sheetData><row r=\"1\"><c r=\"A1\" s=\"1\" t=\"inlineStr\"><is><t>Nine chars</t></is></c></row></sheetData>"
        ),
    );
    let indented = build_xlsx_with_styles_and_sheet(
        &styles,
        &format!(
            "{columns}<sheetData><row r=\"1\"><c r=\"A1\" s=\"2\" t=\"inlineStr\"><is><t>Nine chars</t></is></c></row></sheetData>"
        ),
    );

    let spill_width = |data: &[u8]| -> Option<f64> {
        let (doc, _warnings) = XlsxParser
            .parse(data, &ConvertOptions::default())
            .expect("workbook should parse");
        let tp = get_sheet_page(&doc, 0);
        tp.table.rows[0].cells[0].spill_width
    };

    assert_eq!(
        spill_width(&flush),
        None,
        "the flush line fits its own column"
    );
    assert!(
        spill_width(&indented).is_some(),
        "the same line indented by 12pt no longer fits, so it spills"
    );
}
