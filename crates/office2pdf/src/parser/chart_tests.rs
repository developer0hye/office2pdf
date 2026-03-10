use super::*;

#[test]
fn test_parse_bar_chart() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:chart>
                <c:title><c:tx><c:rich><a:p><a:r><a:t>Sales Data</a:t></a:r></a:p></c:rich></c:tx></c:title>
                <c:plotArea>
                    <c:barChart>
                        <c:ser>
                            <c:idx val="0"/>
                            <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
                            <c:cat>
                                <c:strRef><c:strCache>
                                    <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                                    <c:pt idx="1"><c:v>Q2</c:v></c:pt>
                                    <c:pt idx="2"><c:v>Q3</c:v></c:pt>
                                </c:strCache></c:strRef>
                            </c:cat>
                            <c:val>
                                <c:numRef><c:numCache>
                                    <c:pt idx="0"><c:v>100</c:v></c:pt>
                                    <c:pt idx="1"><c:v>200</c:v></c:pt>
                                    <c:pt idx="2"><c:v>150</c:v></c:pt>
                                </c:numCache></c:numRef>
                            </c:val>
                        </c:ser>
                    </c:barChart>
                </c:plotArea>
            </c:chart>
        </c:chartSpace>"#;

    let chart = parse_chart_xml(xml).unwrap();
    assert_eq!(chart.chart_type, ChartType::Bar);
    assert_eq!(chart.title.as_deref(), Some("Sales Data"));
    assert_eq!(chart.categories, vec!["Q1", "Q2", "Q3"]);
    assert_eq!(chart.series.len(), 1);
    assert_eq!(chart.series[0].name.as_deref(), Some("Revenue"));
    assert_eq!(chart.series[0].values, vec![100.0, 200.0, 150.0]);
}

#[test]
fn test_parse_pie_chart() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
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
                                    <c:pt idx="2"><c:v>Cherry</c:v></c:pt>
                                </c:strLit>
                            </c:cat>
                            <c:val>
                                <c:numLit>
                                    <c:pt idx="0"><c:v>30</c:v></c:pt>
                                    <c:pt idx="1"><c:v>45</c:v></c:pt>
                                    <c:pt idx="2"><c:v>25</c:v></c:pt>
                                </c:numLit>
                            </c:val>
                        </c:ser>
                    </c:pieChart>
                </c:plotArea>
            </c:chart>
        </c:chartSpace>"#;

    let chart = parse_chart_xml(xml).unwrap();
    assert_eq!(chart.chart_type, ChartType::Pie);
    assert!(chart.title.is_none());
    assert_eq!(chart.categories, vec!["Apple", "Banana", "Cherry"]);
    assert_eq!(chart.series[0].values, vec![30.0, 45.0, 25.0]);
}

#[test]
fn test_parse_line_chart_multiple_series() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:chart>
                <c:title><c:tx><c:rich><a:p><a:r><a:t>Trends</a:t></a:r></a:p></c:rich></c:tx></c:title>
                <c:plotArea>
                    <c:lineChart>
                        <c:ser>
                            <c:idx val="0"/>
                            <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series A</c:v></c:pt></c:strCache></c:strRef></c:tx>
                            <c:cat>
                                <c:strRef><c:strCache>
                                    <c:pt idx="0"><c:v>Jan</c:v></c:pt>
                                    <c:pt idx="1"><c:v>Feb</c:v></c:pt>
                                </c:strCache></c:strRef>
                            </c:cat>
                            <c:val>
                                <c:numRef><c:numCache>
                                    <c:pt idx="0"><c:v>10</c:v></c:pt>
                                    <c:pt idx="1"><c:v>20</c:v></c:pt>
                                </c:numCache></c:numRef>
                            </c:val>
                        </c:ser>
                        <c:ser>
                            <c:idx val="1"/>
                            <c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Series B</c:v></c:pt></c:strCache></c:strRef></c:tx>
                            <c:val>
                                <c:numRef><c:numCache>
                                    <c:pt idx="0"><c:v>15</c:v></c:pt>
                                    <c:pt idx="1"><c:v>25</c:v></c:pt>
                                </c:numCache></c:numRef>
                            </c:val>
                        </c:ser>
                    </c:lineChart>
                </c:plotArea>
            </c:chart>
        </c:chartSpace>"#;

    let chart = parse_chart_xml(xml).unwrap();
    assert_eq!(chart.chart_type, ChartType::Line);
    assert_eq!(chart.title.as_deref(), Some("Trends"));
    assert_eq!(chart.categories, vec!["Jan", "Feb"]);
    assert_eq!(chart.series.len(), 2);
    assert_eq!(chart.series[0].name.as_deref(), Some("Series A"));
    assert_eq!(chart.series[0].values, vec![10.0, 20.0]);
    assert_eq!(chart.series[1].name.as_deref(), Some("Series B"));
    assert_eq!(chart.series[1].values, vec![15.0, 25.0]);
}

#[test]
fn test_parse_chart_no_title() {
    let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart>
                <c:plotArea>
                    <c:barChart>
                        <c:ser>
                            <c:idx val="0"/>
                            <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat>
                            <c:val><c:numLit><c:pt idx="0"><c:v>42</c:v></c:pt></c:numLit></c:val>
                        </c:ser>
                    </c:barChart>
                </c:plotArea>
            </c:chart>
        </c:chartSpace>"#;

    let chart = parse_chart_xml(xml).unwrap();
    assert!(chart.title.is_none());
    assert_eq!(chart.categories, vec!["A"]);
    assert_eq!(chart.series[0].values, vec![42.0]);
}

#[test]
fn test_scan_chart_references() {
    let xml = r#"<?xml version="1.0"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <w:body>
                <w:p><w:r><w:t>Hello</w:t></w:r></w:p>
                <w:p>
                    <w:r>
                        <w:drawing>
                            <wp:inline>
                                <a:graphic>
                                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                                        <c:chart r:id="rId4"/>
                                    </a:graphicData>
                                </a:graphic>
                            </wp:inline>
                        </w:drawing>
                    </w:r>
                </w:p>
            </w:body>
        </w:document>"#;

    let refs = scan_chart_references(xml);
    assert_eq!(refs.len(), 1);
    assert_eq!(refs[0].0, 1); // second body child
    assert_eq!(refs[0].1, "rId4");
}

#[test]
fn test_scan_chart_rels() {
    let rels_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
            <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
            <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart2.xml"/>
        </Relationships>"#;

    let rels = scan_chart_rels(rels_xml);
    assert_eq!(rels.len(), 2);
    assert_eq!(rels.get("rId4").unwrap(), "word/charts/chart1.xml");
    assert_eq!(rels.get("rId5").unwrap(), "word/charts/chart2.xml");
}
