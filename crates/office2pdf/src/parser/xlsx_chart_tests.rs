use super::*;

fn build_xlsx_with_chart(cells: &[(&str, &str)], chart_xml: &str) -> Vec<u8> {
    let base = build_xlsx_bytes("Sheet1", cells);

    let reader = std::io::Cursor::new(&base);
    let mut archive = zip::ZipArchive::new(reader).unwrap();

    let mut out_buf = Vec::new();
    {
        let cursor = std::io::Cursor::new(&mut out_buf);
        let mut writer = zip::ZipWriter::new(cursor);

        for i in 0..archive.len() {
            let mut entry = archive.by_index(i).unwrap();
            let options: zip::write::FileOptions =
                zip::write::FileOptions::default().compression_method(entry.compression());
            writer
                .start_file(entry.name().to_string(), options)
                .unwrap();
            std::io::copy(&mut entry, &mut writer).unwrap();
        }

        let options: zip::write::FileOptions = zip::write::FileOptions::default();
        writer.start_file("xl/charts/chart1.xml", options).unwrap();
        use std::io::Write;
        writer.write_all(chart_xml.as_bytes()).unwrap();

        writer.finish().unwrap();
    }

    out_buf
}

fn make_bar_chart_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:chart>
                <c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>
                <c:plotArea>
                    <c:barChart>
                        <c:ser>
                            <c:idx val="0"/>
                            <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
                            <c:cat>
                                <c:strRef><c:strCache>
                                    <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                                    <c:pt idx="1"><c:v>Q2</c:v></c:pt>
                                </c:strCache></c:strRef>
                            </c:cat>
                            <c:val>
                                <c:numRef><c:numCache>
                                    <c:pt idx="0"><c:v>100</c:v></c:pt>
                                    <c:pt idx="1"><c:v>200</c:v></c:pt>
                                </c:numCache></c:numRef>
                            </c:val>
                        </c:ser>
                    </c:barChart>
                </c:plotArea>
            </c:chart>
        </c:chartSpace>"#
        .to_string()
}

#[test]
fn test_xlsx_with_chart_embeds_in_table_page() {
    let data = build_xlsx_with_chart(&[("A1", "Hello")], &make_bar_chart_xml());
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(
        doc.pages.len(),
        1,
        "Expected 1 page (chart embedded in table)"
    );
    assert!(matches!(&doc.pages[0], Page::Table(_)));

    let tp = get_table_page(&doc, 0);
    assert!(!tp.charts.is_empty(), "Expected charts in table page");

    let chart = &tp.charts[0].1;
    assert_eq!(chart.chart_type, ChartType::Bar);
    assert_eq!(chart.title.as_deref(), Some("Sales"));
    assert_eq!(chart.categories, vec!["Q1", "Q2"]);
    assert_eq!(chart.series[0].values, vec![100.0, 200.0]);
}

#[test]
fn test_xlsx_without_chart_no_extra_pages() {
    let data = build_xlsx_bytes("Sheet1", &[("A1", "Hello")]);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(doc.pages.len(), 1);
    assert!(matches!(&doc.pages[0], Page::Table(_)));
}

#[test]
fn test_xlsx_chart_data_is_correct() {
    let chart_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart>
                <c:plotArea>
                    <c:pieChart>
                        <c:ser>
                            <c:idx val="0"/>
                            <c:cat>
                                <c:strLit>
                                    <c:pt idx="0"><c:v>Apple</c:v></c:pt>
                                    <c:pt idx="1"><c:v>Banana</c:v></c:pt>
                                </c:strLit>
                            </c:cat>
                            <c:val>
                                <c:numLit>
                                    <c:pt idx="0"><c:v>60</c:v></c:pt>
                                    <c:pt idx="1"><c:v>40</c:v></c:pt>
                                </c:numLit>
                            </c:val>
                        </c:ser>
                    </c:pieChart>
                </c:plotArea>
            </c:chart>
        </c:chartSpace>"#;

    let data = build_xlsx_with_chart(&[("A1", "Data")], chart_xml);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_table_page(&doc, 0);
    assert!(!tp.charts.is_empty(), "Expected a chart in the table page");
    let chart = &tp.charts[0].1;
    assert_eq!(chart.chart_type, ChartType::Pie);
    assert!(chart.title.is_none());
    assert_eq!(chart.categories, vec!["Apple", "Banana"]);
    assert_eq!(chart.series[0].values, vec![60.0, 40.0]);
}

fn build_xlsx_with_anchored_chart(
    cells: &[(&str, &str)],
    chart_xml: &str,
    anchor_row: u32,
) -> Vec<u8> {
    let base = build_xlsx_bytes("Sheet1", cells);

    let reader = std::io::Cursor::new(&base);
    let mut archive = zip::ZipArchive::new(reader).unwrap();

    let mut workbook_rels_xml = String::new();
    if let Ok(mut entry) = archive.by_name("xl/_rels/workbook.xml.rels") {
        std::io::Read::read_to_string(&mut entry, &mut workbook_rels_xml).unwrap();
    }
    let sheet_target = workbook_rels_xml
        .split("Target=\"")
        .filter_map(|segment| {
            let end = segment.find('"')?;
            let target = &segment[..end];
            if target.contains("worksheets/") {
                Some(target.to_string())
            } else {
                None
            }
        })
        .next()
        .unwrap_or_else(|| "worksheets/sheet1.xml".to_string());

    let sheet_filename = sheet_target.rsplit('/').next().unwrap();
    let sheet_rels_path = format!("xl/worksheets/_rels/{sheet_filename}.rels");

    let mut out_buf = Vec::new();
    {
        let cursor = std::io::Cursor::new(&mut out_buf);
        let mut writer = zip::ZipWriter::new(cursor);
        let options: zip::write::FileOptions = zip::write::FileOptions::default();

        for i in 0..archive.len() {
            let mut entry = archive.by_index(i).unwrap();
            let name = entry.name().to_string();
            writer.start_file(name, options).unwrap();
            std::io::copy(&mut entry, &mut writer).unwrap();
        }

        writer.start_file(&sheet_rels_path, options).unwrap();
        use std::io::Write;
        writer
            .write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>"#,
            )
            .unwrap();

        writer
            .start_file("xl/drawings/drawing1.xml", options)
            .unwrap();
        let drawing_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>2</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>{anchor_row}</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>8</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>{}</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="1" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrmPr/>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#,
            anchor_row + 15
        );
        writer.write_all(drawing_xml.as_bytes()).unwrap();

        writer
            .start_file("xl/drawings/_rels/drawing1.xml.rels", options)
            .unwrap();
        writer
            .write_all(
                br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>"#,
            )
            .unwrap();

        writer.start_file("xl/charts/chart1.xml", options).unwrap();
        writer.write_all(chart_xml.as_bytes()).unwrap();

        writer.finish().unwrap();
    }

    out_buf
}

#[test]
fn test_xlsx_chart_anchored_at_row_5() {
    let cells: Vec<(&str, &str)> = (1..=10)
        .map(|row| {
            let coord: &str = Box::leak(format!("A{row}").into_boxed_str());
            let value: &str = Box::leak(format!("Row {row}").into_boxed_str());
            (coord, value)
        })
        .collect();

    let data = build_xlsx_with_anchored_chart(&cells, &make_bar_chart_xml(), 5);
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    assert_eq!(
        doc.pages.len(),
        1,
        "Chart with anchor should be embedded in table page, not separate"
    );

    let tp = get_table_page(&doc, 0);
    assert_eq!(tp.charts.len(), 1, "Expected 1 anchored chart");
    assert_eq!(tp.charts[0].0, 5, "Chart should be anchored at row 5");
    assert_eq!(tp.charts[0].1.chart_type, ChartType::Bar);
    assert_eq!(tp.charts[0].1.title.as_deref(), Some("Sales"));
}

#[test]
fn test_xlsx_chart_without_anchor_falls_back_to_end() {
    let data = build_xlsx_with_chart(&[("A1", "Hello")], &make_bar_chart_xml());
    let parser = XlsxParser;
    let (doc, _warnings) = parser.parse(&data, &ConvertOptions::default()).unwrap();

    let tp = get_table_page(&doc, 0);
    assert!(
        !tp.charts.is_empty(),
        "Unanchored chart should still be embedded in table page"
    );
    assert_eq!(
        tp.charts[0].0,
        u32::MAX,
        "Unanchored chart should have sentinel row"
    );
}
