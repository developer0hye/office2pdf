use super::*;

#[test]
fn test_codegen_chart_bar_visual_bars() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Chart {
        chart_type: ChartType::Bar,
        title: Some("Sales Report".to_string()),
        categories: vec!["Q1".to_string(), "Q2".to_string()],
        series: vec![ChartSeries {
            name: Some("Revenue".to_string()),
            values: vec![100.0, 250.0],
        }],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("stroke:"),
        "Expected bordered box, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Sales Report"),
        "Expected chart title in header, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Bar Chart"),
        "Expected chart type label, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("box(") || output.source.contains("#box("),
        "Expected visual bar boxes for bar chart, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Q1"),
        "Expected category label, got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_chart_pie_percentages() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Chart {
        chart_type: ChartType::Pie,
        title: Some("Market Share".to_string()),
        categories: vec!["A".to_string(), "B".to_string()],
        series: vec![ChartSeries {
            name: None,
            values: vec![60.0, 40.0],
        }],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("Pie Chart"),
        "Expected pie chart label, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("60") && output.source.contains("%"),
        "Expected percentage in pie chart, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("40") && output.source.contains("%"),
        "Expected percentage in pie chart, got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_chart_line_trend_indicators() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Chart {
        chart_type: ChartType::Line,
        title: Some("Trends".to_string()),
        categories: vec!["Jan".to_string(), "Feb".to_string(), "Mar".to_string()],
        series: vec![ChartSeries {
            name: Some("Sales".to_string()),
            values: vec![10.0, 20.0, 15.0],
        }],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("Line Chart"),
        "Expected line chart label, got:\n{}",
        output.source
    );
    let has_trend =
        output.source.contains('↑') || output.source.contains('↓') || output.source.contains('→');
    assert!(
        has_trend,
        "Expected trend indicators in line chart, got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_chart_empty_series() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Chart {
        chart_type: ChartType::Line,
        title: Some("Empty".to_string()),
        categories: vec![],
        series: vec![],
    })])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("Line Chart"),
        "Expected line chart label, got:\n{}",
        output.source
    );
}

fn sa_node(text: &str, depth: usize) -> SmartArtNode {
    SmartArtNode {
        text: text.to_string(),
        depth,
    }
}

#[test]
fn test_smartart_codegen_flat_numbered_steps() {
    let doc = make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![FixedElement {
            x: 72.0,
            y: 100.0,
            width: 400.0,
            height: 300.0,
            kind: FixedElementKind::SmartArt(SmartArt {
                items: vec![
                    sa_node("Step 1", 0),
                    sa_node("Step 2", 0),
                    sa_node("Step 3", 0),
                ],
            }),
        }],
    )]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("stroke:"),
        "Expected bordered box, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("SmartArt Diagram"),
        "Expected SmartArt header, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Step 1"),
        "Expected Step 1, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Step 2"),
        "Expected Step 2, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Step 3"),
        "Expected Step 3, got:\n{}",
        output.source
    );
}

#[test]
fn test_smartart_codegen_hierarchy_indented_tree() {
    let doc = make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![FixedElement {
            x: 72.0,
            y: 100.0,
            width: 400.0,
            height: 300.0,
            kind: FixedElementKind::SmartArt(SmartArt {
                items: vec![
                    sa_node("CEO", 0),
                    sa_node("VP Engineering", 1),
                    sa_node("VP Sales", 1),
                    sa_node("Dev Lead", 2),
                ],
            }),
        }],
    )]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("CEO"),
        "Expected CEO, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("pad"),
        "Expected indented items for hierarchy, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("VP Engineering"),
        "Expected VP Engineering, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Dev Lead"),
        "Expected Dev Lead, got:\n{}",
        output.source
    );
}

#[test]
fn test_smartart_codegen_empty_items() {
    let doc = make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![FixedElement {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 100.0,
            kind: FixedElementKind::SmartArt(SmartArt { items: vec![] }),
        }],
    )]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("SmartArt Diagram"),
        "Expected SmartArt header even for empty SmartArt"
    );
}

#[test]
fn test_smartart_codegen_special_chars() {
    let doc = make_doc(vec![make_fixed_page(
        720.0,
        540.0,
        vec![FixedElement {
            x: 0.0,
            y: 0.0,
            width: 200.0,
            height: 100.0,
            kind: FixedElementKind::SmartArt(SmartArt {
                items: vec![sa_node("Item #1", 0), sa_node("Price $10", 0)],
            }),
        }],
    )]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains(r"\#"),
        "Expected escaped #, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains(r"\$"),
        "Expected escaped $, got:\n{}",
        output.source
    );
}
