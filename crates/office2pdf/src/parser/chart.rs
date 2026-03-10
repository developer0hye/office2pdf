//! Chart XML parser for DOCX embedded charts.
//!
//! Parses chart*.xml files from DOCX ZIP archives and extracts chart type,
//! title, category labels, and series data into IR `Chart` structs.

use quick_xml::Reader;
use quick_xml::events::Event;

use super::xml_util;
use crate::ir::{Chart, ChartSeries, ChartType};

/// Mapping from XML chart element tag names to their corresponding `ChartType`.
/// Both 2-D and 3-D variants map to the same logical type.
const CHART_TAG_TYPES: &[(&[u8], ChartType)] = &[
    (b"barChart", ChartType::Bar),
    (b"bar3DChart", ChartType::Bar),
    (b"lineChart", ChartType::Line),
    (b"line3DChart", ChartType::Line),
    (b"pieChart", ChartType::Pie),
    (b"pie3DChart", ChartType::Pie),
    (b"areaChart", ChartType::Area),
    (b"scatterChart", ChartType::Scatter),
];

/// Look up a tag name in [`CHART_TAG_TYPES`] and return the matching `ChartType`.
fn chart_type_for_tag(tag: &[u8]) -> Option<ChartType> {
    CHART_TAG_TYPES
        .iter()
        .find(|(name, _)| *name == tag)
        .map(|(_, ct)| ct.clone())
}

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
                let tag: &[u8] = local.as_ref();
                if tag == b"title" && title.is_none() {
                    title = parse_chart_title(&mut reader);
                } else if let Some(ct) = chart_type_for_tag(tag) {
                    chart_type = Some(ct);
                    parse_chart_series(&mut reader, tag, &mut categories, &mut series);
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
                        xml_util::skip_element(reader, b"xVal");
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
