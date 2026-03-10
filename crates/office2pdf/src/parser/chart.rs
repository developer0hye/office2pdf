//! Chart XML parser for DOCX embedded charts.
//!
//! Parses chart*.xml files from DOCX ZIP archives and extracts chart type,
//! title, category labels, and series data into IR `Chart` structs.

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::ir::{Chart, ChartSeries, ChartType};

/// Parse a chart XML file (e.g., `word/charts/chart1.xml`) into a `Chart` IR.
pub(crate) fn parse_chart_xml(xml: &str) -> Option<Chart> {
    let mut reader = Reader::from_str(xml);
    let mut chart_type = None;
    let mut title = None;
    let mut categories: Vec<String> = Vec::new();
    let mut series: Vec<ChartSeries> = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"title" if title.is_none() => {
                        title = parse_chart_title(&mut reader);
                    }
                    b"barChart" => {
                        chart_type = Some(ChartType::Bar);
                        parse_chart_series(&mut reader, b"barChart", &mut categories, &mut series);
                    }
                    b"bar3DChart" => {
                        chart_type = Some(ChartType::Bar);
                        parse_chart_series(
                            &mut reader,
                            b"bar3DChart",
                            &mut categories,
                            &mut series,
                        );
                    }
                    b"lineChart" => {
                        chart_type = Some(ChartType::Line);
                        parse_chart_series(&mut reader, b"lineChart", &mut categories, &mut series);
                    }
                    b"line3DChart" => {
                        chart_type = Some(ChartType::Line);
                        parse_chart_series(
                            &mut reader,
                            b"line3DChart",
                            &mut categories,
                            &mut series,
                        );
                    }
                    b"pieChart" => {
                        chart_type = Some(ChartType::Pie);
                        parse_chart_series(&mut reader, b"pieChart", &mut categories, &mut series);
                    }
                    b"pie3DChart" => {
                        chart_type = Some(ChartType::Pie);
                        parse_chart_series(
                            &mut reader,
                            b"pie3DChart",
                            &mut categories,
                            &mut series,
                        );
                    }
                    b"areaChart" => {
                        chart_type = Some(ChartType::Area);
                        parse_chart_series(&mut reader, b"areaChart", &mut categories, &mut series);
                    }
                    b"scatterChart" => {
                        chart_type = Some(ChartType::Scatter);
                        parse_chart_series(
                            &mut reader,
                            b"scatterChart",
                            &mut categories,
                            &mut series,
                        );
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    let chart_type = chart_type?;
    Some(Chart {
        chart_type,
        title,
        categories,
        series,
    })
}

/// Parse the chart title text from `<c:title>`.
fn parse_chart_title(reader: &mut Reader<&[u8]>) -> Option<String> {
    let mut text = String::new();
    let mut in_t = false;
    let mut depth = 1u32;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"title" {
                    depth += 1;
                } else if local.as_ref() == b"t" {
                    in_t = true;
                }
            }
            Ok(Event::Text(ref t)) if in_t => {
                if let Ok(s) = t.xml_content() {
                    text.push_str(s.as_ref());
                }
            }
            Ok(Event::End(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"t" {
                    in_t = false;
                } else if local.as_ref() == b"title" {
                    depth -= 1;
                    if depth == 0 {
                        break;
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let trimmed = text.trim().to_string();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed)
    }
}

/// Parse series data from within a chart type element (e.g., `<c:barChart>`).
fn parse_chart_series(
    reader: &mut Reader<&[u8]>,
    end_tag: &[u8],
    categories: &mut Vec<String>,
    series: &mut Vec<ChartSeries>,
) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"ser" {
                    let (ser, cats) = parse_single_series(reader);
                    // Use categories from first series that has them
                    if categories.is_empty() && !cats.is_empty() {
                        *categories = cats;
                    }
                    series.push(ser);
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == end_tag => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

/// Parse a single `<c:ser>` element and return the series data + category labels.
fn parse_single_series(reader: &mut Reader<&[u8]>) -> (ChartSeries, Vec<String>) {
    let mut name = None;
    let mut values = Vec::new();
    let mut categories = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"tx" => name = parse_series_text(reader),
                b"cat" => categories = parse_category_data(reader),
                b"val" | b"yVal" => values = parse_value_data(reader),
                b"xVal" => {
                    // For scatter charts, xVal contains category-like data
                    if categories.is_empty() {
                        categories = parse_category_data(reader);
                    } else {
                        skip_element(reader, b"xVal");
                    }
                }
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"ser" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    (ChartSeries { name, values }, categories)
}

/// Parse series name from `<c:tx>`.
fn parse_series_text(reader: &mut Reader<&[u8]>) -> Option<String> {
    let mut text = String::new();
    let mut in_v = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"v" {
                    in_v = true;
                }
            }
            Ok(Event::Text(ref t)) if in_v => {
                if let Ok(s) = t.xml_content() {
                    text.push_str(s.as_ref());
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"v" => in_v = false,
                b"tx" => break,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let trimmed = text.trim().to_string();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed)
    }
}

/// Parse category labels from `<c:cat>` (either `<c:strRef>` or `<c:strLit>`).
fn parse_category_data(reader: &mut Reader<&[u8]>) -> Vec<String> {
    let mut categories = Vec::new();
    let mut in_v = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"v" {
                    in_v = true;
                }
            }
            Ok(Event::Text(ref t)) if in_v => {
                if let Ok(s) = t.xml_content() {
                    categories.push(s.as_ref().to_string());
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"v" => in_v = false,
                b"cat" | b"xVal" => break,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    categories
}

/// Parse numeric values from `<c:val>` or `<c:yVal>`.
fn parse_value_data(reader: &mut Reader<&[u8]>) -> Vec<f64> {
    let mut values = Vec::new();
    let mut in_v = false;
    let mut current_text = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"v" {
                    in_v = true;
                    current_text.clear();
                }
            }
            Ok(Event::Text(ref t)) if in_v => {
                if let Ok(s) = t.xml_content() {
                    current_text.push_str(s.as_ref());
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"v" => {
                    in_v = false;
                    if let Ok(v) = current_text.trim().parse::<f64>() {
                        values.push(v);
                    }
                }
                b"val" | b"yVal" => break,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    values
}

fn skip_element(reader: &mut Reader<&[u8]>, end_tag: &[u8]) {
    let mut depth = 1u32;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == end_tag {
                    depth += 1;
                }
            }
            Ok(Event::End(ref e)) => {
                if e.local_name().as_ref() == end_tag {
                    depth -= 1;
                    if depth == 0 {
                        return;
                    }
                }
            }
            Ok(Event::Eof) | Err(_) => return,
            _ => {}
        }
    }
}

/// Scan document.xml for chart relationship IDs.
///
/// Returns `(body_child_index, relationship_id)` tuples for each chart reference
/// found in drawing elements.
pub(crate) fn scan_chart_references(xml: &str) -> Vec<(usize, String)> {
    let mut results = Vec::new();
    let mut reader = Reader::from_str(xml);

    let mut in_body = false;
    let mut body_child_index: usize = 0;
    let mut depth_in_body: u32 = 0;
    let mut in_graphic_data = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();

                if name == b"body" {
                    in_body = true;
                    depth_in_body = 0;
                    body_child_index = 0;
                    continue;
                }

                if in_body {
                    depth_in_body += 1;
                }

                if name == b"graphicData" {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"uri"
                            && let Ok(val) = attr.unescape_value()
                            && val.contains("chart")
                        {
                            in_graphic_data = true;
                        }
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                let name = local.as_ref();

                if in_body {
                    depth_in_body += 1;
                    // Empty elements open and close immediately
                    depth_in_body -= 1;
                }

                if in_graphic_data && name == b"chart" {
                    for attr in e.attributes().flatten() {
                        if attr.key.local_name().as_ref() == b"id"
                            && let Ok(val) = attr.unescape_value()
                        {
                            results.push((body_child_index, val.to_string()));
                        }
                    }
                }

                // Empty graphicData can't contain a chart child element, skip
            }
            Ok(Event::End(ref e)) => {
                let name = e.local_name();
                if name.as_ref() == b"body" {
                    in_body = false;
                } else if name.as_ref() == b"graphicData" {
                    in_graphic_data = false;
                } else if in_body && depth_in_body > 0 {
                    depth_in_body -= 1;
                    if depth_in_body == 0 {
                        body_child_index += 1;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    results
}

/// Scan `word/_rels/document.xml.rels` for chart relationship targets.
///
/// Returns a map from relationship ID to chart file path (e.g., "rId4" → "word/charts/chart1.xml").
pub(crate) fn scan_chart_rels(rels_xml: &str) -> std::collections::HashMap<String, String> {
    let mut map = std::collections::HashMap::new();
    let mut reader = Reader::from_str(rels_xml);

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let mut id = None;
                    let mut target = None;
                    let mut is_chart = false;

                    for attr in e.attributes().flatten() {
                        match attr.key.local_name().as_ref() {
                            b"Id" => {
                                if let Ok(v) = attr.unescape_value() {
                                    id = Some(v.to_string());
                                }
                            }
                            b"Target" => {
                                if let Ok(v) = attr.unescape_value() {
                                    target = Some(v.to_string());
                                }
                            }
                            b"Type" => {
                                if let Ok(v) = attr.unescape_value()
                                    && v.contains("chart")
                                {
                                    is_chart = true;
                                }
                            }
                            _ => {}
                        }
                    }

                    if is_chart && let (Some(id), Some(target)) = (id, target) {
                        // Target is relative to word/ directory
                        let full_path = if let Some(stripped) = target.strip_prefix('/') {
                            stripped.to_string()
                        } else {
                            format!("word/{target}")
                        };
                        map.insert(id, full_path);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    map
}

#[cfg(test)]
#[path = "chart_tests.rs"]
mod tests;
