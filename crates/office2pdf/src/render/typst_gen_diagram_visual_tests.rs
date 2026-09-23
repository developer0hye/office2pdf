use super::*;
use crate::ir::DataLabels;
use crate::ir::MarkerSymbol;
use crate::ir::{ChartAreaFill, ChartAreaOutline};
use crate::render::typst_gen::diagrams::{
    CHART_AREA_OUTLINE, CHART_AUTOMATIC_LINE, CHART_DEFAULT_TEXT_PT, GAP, LABEL_W, LEGEND_ENTRY_W,
    LEGEND_KEY_BASELINE_PT, LEGEND_KEY_LEN_PT, PPTX_LEGEND_KEY_EM,
    PPTX_LEGEND_KEY_LABEL_GAP_KEY_SHARE, PPTX_LEGEND_KEY_LABEL_GAP_PT,
    PPTX_RIGHT_LEGEND_Y_SHIFT_EM, PPTX_RIGHT_LEGEND_Y_SHIFT_PT, ROW, SERIES_LINE_PT,
    SERIES_MARKER_SIZE_PT, TICK_GAP, axis_plot_rect, chart_area_title_h, chart_category_band_pt,
    chart_category_gutter_pt, chart_category_rotated_label_x, chart_category_rotated_label_y,
    chart_face_line_metrics_em, chart_text_advance_em, chart_tick_band_pt,
    excel_legend_trailing_gutter_pt, legend_marker_cap_pt, powerpoint_right_legend_y_shift,
    pptx_column_data_label_seat_pt,
};

#[test]
fn test_codegen_chart_bar_visual_bars() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Box::new(Chart {
        chart_type: ChartType::Bar,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Sales Report".to_string()),
        categories: vec!["Q1".to_string(), "Q2".to_string()],
        series: vec![ChartSeries {
            name: Some("Revenue".to_string()),
            values: vec![100.0, 250.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }))])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("Sales Report"),
        "Expected chart title, got:\n{}",
        output.source
    );
    // Axis-scaled bar chart: series-name area title, rect bars, tick labels,
    // and gridlines (no raw "Bar Chart" placeholder or bordered box).
    assert!(
        output.source.contains("Revenue"),
        "Expected series-name area title, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("rect(width:"),
        "Expected axis-scaled bar rects, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("line(end:"),
        "Expected axis gridlines, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Q1"),
        "Expected category label, got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_chart_axis_ticks_and_no_raw_floats() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Box::new(Chart {
        chart_type: ChartType::Bar,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("My Bar Chart".to_string()),
        categories: vec!["1st Qtr".to_string(), "2nd Qtr".to_string()],
        series: vec![ChartSeries {
            name: Some("Sales".to_string()),
            values: vec![8.200000000000001, 3.2],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }))])]);

    let output = generate_typst(&doc).unwrap();
    // Bars carry no in-plot value labels (like PowerPoint), so the raw float
    // never reaches the output.
    assert!(
        !output.source.contains("8.200000000000001"),
        "raw float must not leak; got:\n{}",
        output.source
    );
    // Nice axis for max 8.2 → ticks 0,1,…,9.
    for tick in ["[0]", "[1]", "[9]"] {
        assert!(
            output.source.contains(tick),
            "expected axis tick {tick}; got:\n{}",
            output.source
        );
    }
}

#[test]
fn test_codegen_chart_pie_draws_a_pie() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Box::new(Chart {
        chart_type: ChartType::Pie,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Market Share".to_string()),
        categories: vec!["A".to_string(), "B".to_string()],
        series: vec![ChartSeries {
            name: None,
            values: vec![60.0, 40.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }))])]);

    let output = generate_typst(&doc).unwrap();

    // A pie is a pie, not the `Slice | Value | %` table it used to be (#533).
    assert!(
        output.source.contains("Market Share"),
        "Expected chart title, got:\n{}",
        output.source
    );
    assert_eq!(
        output.source.matches("curve(fill:").count(),
        2,
        "one wedge per slice, got:\n{}",
        output.source
    );
    for category in ["A", "B"] {
        assert!(
            output.source.contains(category),
            "Expected {category} in the legend, got:\n{}",
            output.source
        );
    }
}

#[test]
fn test_codegen_chart_line_trend_indicators() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Box::new(Chart {
        chart_type: ChartType::Line,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Trends".to_string()),
        categories: vec!["Jan".to_string(), "Feb".to_string(), "Mar".to_string()],
        series: vec![ChartSeries {
            name: Some("Sales".to_string()),
            values: vec![10.0, 20.0, 15.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }))])]);

    let output = generate_typst(&doc).unwrap();
    // Multi-point line charts now render as an axis-scaled polyline plot
    // (not a trend-indicator table).
    assert!(
        output.source.contains("Trends"),
        "Expected chart title, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("curve(stroke:"),
        "Expected polyline curve for the line chart, got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("Sales"),
        "Expected series name in legend, got:\n{}",
        output.source
    );
}

#[test]
fn test_codegen_chart_empty_series() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Box::new(Chart {
        chart_type: ChartType::Line,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Empty".to_string()),
        categories: vec![],
        series: vec![],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }))])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("Line Chart"),
        "Expected line chart label, got:\n{}",
        output.source
    );
}

/// Render `doc` and return the visible text of each PDF page separately.
fn page_texts(doc: &Document) -> Vec<String> {
    let pdf: Vec<u8> = crate::render_document(doc).unwrap();
    pdf_extract::extract_text_from_mem_by_pages(&pdf).unwrap()
}

/// Fill most of a page so the chart that follows cannot fit in what is left.
fn page_filler(lines: usize) -> Vec<Block> {
    (1..=lines)
        .map(|line| {
            make_paragraph(&format!(
                "Line {line} of the quarterly commentary preceding the chart."
            ))
        })
        .collect()
}

/// Report the index of the single page carrying `marker`, or panic with the
/// page breakdown when it is missing or duplicated.
fn page_holding(pages: &[String], marker: &str) -> usize {
    let hits: Vec<usize> = pages
        .iter()
        .enumerate()
        .filter(|(_, text)| text.contains(marker))
        .map(|(index, _)| index)
        .collect();
    assert_eq!(
        hits.len(),
        1,
        "expected {marker:?} on exactly one page, found it on {hits:?}; pages:\n{pages:#?}"
    );
    hits[0]
}

#[test]
fn an_axis_chart_that_does_not_fit_moves_to_the_next_page_whole() {
    let mut content: Vec<Block> = page_filler(30);
    content.push(Block::Chart(Box::new(Chart {
        chart_type: ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Quarterly Units Shipped".to_string()),
        categories: vec![
            "Northlake".to_string(),
            "Eastport".to_string(),
            "Southgate".to_string(),
        ],
        series: vec![ChartSeries {
            name: Some("Units".to_string()),
            values: vec![23334.0, 8331.0, 2727.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    })));
    let doc = make_doc(vec![make_flow_page(content)]);

    let pages = page_texts(&doc);

    // Excel treats a chart as one floating graphic: the title and the plot it
    // labels never land on opposite sides of a page break.
    assert_eq!(
        page_holding(&pages, "Quarterly Units Shipped"),
        page_holding(&pages, "Southgate"),
        "chart title and category labels split across pages; pages:\n{pages:#?}"
    );
}

#[test]
fn a_bordered_chart_box_that_does_not_fit_moves_to_the_next_page_whole() {
    // The pie fallback draws a bordered box; a breakable one closes with a
    // bottom border at the page end and re-opens with a fresh top border, so
    // one chart reads as two.
    let mut content: Vec<Block> = page_filler(30);
    content.push(Block::Chart(Box::new(Chart {
        chart_type: ChartType::Pie,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Fixture Documents by Format".to_string()),
        categories: vec!["DOCX".to_string(), "PPTX".to_string(), "XLSX".to_string()],
        series: vec![ChartSeries {
            name: None,
            values: vec![115.0, 92.0, 138.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    })));
    let doc = make_doc(vec![make_flow_page(content)]);

    let pages = page_texts(&doc);

    assert_eq!(
        page_holding(&pages, "Fixture Documents by Format"),
        page_holding(&pages, "XLSX"),
        "chart box split across pages; pages:\n{pages:#?}"
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

#[test]
fn test_codegen_chart_line_plot() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Box::new(Chart {
        chart_type: ChartType::Line,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["1".to_string(), "2".to_string(), "3".to_string()],
        series: vec![
            ChartSeries {
                name: Some("A".to_string()),
                values: vec![1.0, 2.0, 3.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
            ChartSeries {
                name: Some("B".to_string()),
                values: vec![10.0, 9.0, 14.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
        ],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }))])]);

    let output = generate_typst(&doc).unwrap();
    assert!(
        output.source.contains("curve(stroke:"),
        "line chart must draw polyline curves; got:\n{}",
        output.source
    );
    assert!(
        output.source.contains("line(end:"),
        "line chart must draw axis gridlines; got:\n{}",
        output.source
    );
    // Both series names appear in the legend.
    assert!(output.source.contains("[A]") && output.source.contains("[B]"));
    // Category labels 1..3 present.
    assert!(output.source.contains("[1]") && output.source.contains("[3]"));
}

#[test]
fn a_chart_too_tall_for_a_page_still_breaks_rather_than_overflowing() {
    // Keeping an over-tall chart atomic does not move it to the next page —
    // Typst runs it off the page edge and the overflow is never drawn. Such a
    // chart stays breakable so every row survives.
    let categories: Vec<String> = (1..=60).map(|i| format!("Category{i:03}")).collect();
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Box::new(Chart {
        chart_type: ChartType::Scatter,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Sixty Sample Sites".to_string()),
        categories: categories.clone(),
        series: vec![ChartSeries {
            name: Some("Reading".to_string()),
            values: (1..=60).map(|value| value as f64).collect(),
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }))])]);

    let pages = page_texts(&doc);

    assert!(
        pages.len() > 1,
        "a 60-row chart cannot fit one page; it must break instead of overflowing: {pages:#?}"
    );
    for category in [&categories[0], &categories[59]] {
        assert!(
            pages.iter().any(|page| page.contains(category)),
            "{category} was dropped; pages:\n{pages:#?}"
        );
    }
}

// ----- Stacked grouping (issue #545) -----

/// The introduction deck's slide 17 chart: three formats, four support areas
/// stacked per format. Stack totals are 9, 9, and 6.
///
/// The band layout is the one that slide's `<c:barChart>` declares, so the
/// geometry these tests read is the geometry the fixture asks for.
fn stacked_support_chart(grouping: ChartGrouping) -> Chart {
    Chart {
        chart_type: ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Supported elements by format".to_string()),
        categories: vec!["DOCX".to_string(), "PPTX".to_string(), "XLSX".to_string()],
        series: vec![
            ChartSeries {
                name: Some("Text".to_string()),
                values: vec![4.0, 2.0, 2.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
            ChartSeries {
                name: Some("Tables".to_string()),
                values: vec![1.0, 1.0, 1.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
            ChartSeries {
                name: Some("Graphics".to_string()),
                values: vec![2.0, 4.0, 0.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
            ChartSeries {
                name: Some("Structure".to_string()),
                values: vec![2.0, 2.0, 3.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
        ],
        grouping,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout {
            gap_width_percent: 90.0,
            overlap_percent: 100.0,
        },
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

fn chart_source(chart: Chart) -> String {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Box::new(chart))])]);
    generate_typst(&doc).unwrap().source
}

fn framed_chart_source(chart: &Chart, width: f64, height: f64) -> String {
    let mut source = String::new();
    generate_chart_in(&mut source, chart, Some((width, height)));
    source
}

/// The axis tick labels the generator emitted, in the order written.
fn emitted_axis_ticks(source: &str) -> Vec<f64> {
    emitted_axis_ticks_at_size(source, CHART_DEFAULT_TEXT_PT)
}

fn emitted_axis_ticks_at_size(source: &str, size_pt: f64) -> Vec<f64> {
    let marker: String = format!("text(size: {}pt)[", format_f64(size_pt));
    source
        .lines()
        .filter(|line| line.contains("#place") && line.contains(&marker))
        .filter_map(|line| {
            let after = line.rsplit_once(marker.as_str())?.1;
            // A minus sign reaches the markup escaped — Typst reads a bare `-`
            // before digits as its own Unicode minus — so the label is read
            // back through the same escape (issue #1184).
            after
                .split_once(']')?
                .0
                .replace('\\', "")
                .parse::<f64>()
                .ok()
        })
        .collect()
}

#[test]
fn a_stacked_column_scales_its_axis_to_the_stack_total() {
    // Rendering a stacked chart clustered does not merely look different, it
    // reports different numbers: the axis topped out at the largest single
    // segment (4) instead of the largest stack (9), so no bar could be read as
    // its category's total (#545).
    let ticks = emitted_axis_ticks(&chart_source(stacked_support_chart(ChartGrouping::Stacked)));

    let axis_max: f64 = ticks.iter().copied().fold(0.0, f64::max);
    assert!(
        axis_max >= 9.0,
        "the axis must reach the tallest stack of 9, got {axis_max} from {ticks:?}"
    );
}

#[test]
fn a_clustered_column_still_scales_to_the_largest_segment() {
    // Control: the same data clustered keeps today's axis, so the stacked
    // branch cannot be a blanket change to axis scaling.
    let ticks = emitted_axis_ticks(&chart_source(stacked_support_chart(
        ChartGrouping::Clustered,
    )));

    let axis_max: f64 = ticks.iter().copied().fold(0.0, f64::max);
    assert!(
        (4.0..9.0).contains(&axis_max),
        "a clustered axis covers the largest segment of 4, got {axis_max} from {ticks:?}"
    );
}

#[test]
fn a_stacked_column_draws_one_bar_per_category() {
    // Four series over three categories: clustered draws 12 rects, stacked
    // draws 12 segments too, but they share three x positions instead of
    // spreading across twelve. It is the deck's `<c:overlap val="100"/>` that
    // puts them on one x — grouping alone does not, as
    // `a_stacked_category_divides_its_band_by_the_same_law_a_clustered_one_does`
    // pins against PowerPoint.
    let source = chart_source(stacked_support_chart(ChartGrouping::Stacked));
    let x_positions: std::collections::BTreeSet<String> = source
        .lines()
        .filter(|line| line.contains("rect(width:"))
        .filter_map(|line| {
            let after = line.split_once("dx: ")?.1;
            Some(after.split_once("pt")?.0.to_string())
        })
        .collect();

    assert_eq!(
        x_positions.len(),
        3,
        "a stacked column puts every series at its category's x, got {x_positions:?}"
    );
}

#[test]
fn a_percent_stacked_column_normalises_every_stack() {
    // XLSX totals 6 against DOCX's 9, but both fill the plot completely.
    let source = chart_source(stacked_support_chart(ChartGrouping::PercentStacked));
    let ticks = emitted_axis_ticks(&source);

    let axis_max: f64 = ticks.iter().copied().fold(0.0, f64::max);
    assert!(
        (axis_max - 100.0).abs() < f64::EPSILON,
        "a percent-stacked axis runs to 100, got {axis_max} from {ticks:?}"
    );
}

// ----- Legend position (issue #546) -----

fn legend_chart(position: LegendPosition) -> Chart {
    Chart {
        chart_type: ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Supported elements by format".to_string()),
        categories: vec!["DOCX".to_string(), "PPTX".to_string(), "XLSX".to_string()],
        series: vec![
            ChartSeries {
                name: Some("Text".to_string()),
                values: vec![4.0, 2.0, 2.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
            ChartSeries {
                name: Some("Tables".to_string()),
                values: vec![1.0, 1.0, 1.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
        ],
        grouping: ChartGrouping::Stacked,
        legend_position: position,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

/// The `(x, y)` of every legend entry the generator placed, in emit order.
fn emitted_legend_entries(source: &str) -> Vec<(f64, f64)> {
    source
        .lines()
        .filter(|line| line.contains("box[#box(width: 9pt, height: 9pt"))
        .filter_map(|line| {
            let x = line
                .split_once("dx: ")?
                .1
                .split_once("pt")?
                .0
                .parse()
                .ok()?;
            let y = line
                .split_once("dy: ")?
                .1
                .split_once("pt")?
                .0
                .parse()
                .ok()?;
            Some((x, y))
        })
        .collect()
}

#[test]
fn a_bottom_legend_lays_its_entries_out_side_by_side() {
    // A legend runs along the edge it sits on, so `val="b"` must spread the
    // entries left to right instead of stacking them (#546).
    let entries = emitted_legend_entries(&chart_source(legend_chart(LegendPosition::Bottom)));

    assert_eq!(
        entries.len(),
        2,
        "expected one entry per series: {entries:?}"
    );
    assert!(
        entries[1].0 > entries[0].0,
        "entries must advance across the page: {entries:?}"
    );
    assert!(
        (entries[0].1 - entries[1].1).abs() < 0.01,
        "entries must share one row: {entries:?}"
    );
}

#[test]
fn a_right_legend_still_stacks_its_entries() {
    // Control: the default position keeps the vertical stack, so the bottom
    // branch cannot be a blanket change to legend layout.
    let entries = emitted_legend_entries(&chart_source(legend_chart(LegendPosition::Right)));

    assert_eq!(entries.len(), 2);
    assert!(
        (entries[0].0 - entries[1].0).abs() < 0.01,
        "entries must share one column: {entries:?}"
    );
    assert!(
        entries[1].1 > entries[0].1,
        "entries must advance down the page: {entries:?}"
    );
}

#[test]
fn every_legend_family_uses_the_legends_own_run_properties() {
    let style = crate::ir::ChartTextStyle {
        size_pt: Some(17.0),
        bold: Some(true),
        letter_spacing_hundredths: Some(125),
        color: Some(Color::new(0xC0, 0x2A, 0x7A)),
        ellipsis_overflow: false,
    };
    for (chart_type, label) in [
        (ChartType::Column, "Text"),
        (ChartType::Line, "Text"),
        (
            ChartType::Other(crate::ir::RADAR_CHART_LABEL.to_string()),
            "Text",
        ),
        (ChartType::Pie, "DOCX"),
    ] {
        let mut chart = legend_chart(LegendPosition::Bottom);
        chart.chart_type = chart_type.clone();
        chart.legend_text_style = style;
        let source = framed_chart_source(&chart, 480.0, 320.0);
        let entry = source
            .lines()
            .find(|line| line.contains(&format!("[{label}]")))
            .unwrap_or_else(|| panic!("{chart_type:?} emits a {label} legend entry"));

        assert!(
            entry.contains(
                "#text(size: 17pt, weight: \"bold\", fill: rgb(192, 42, 122), tracking: 1.25pt, ligatures: false, kerning: false)"
            ),
            "{chart_type:?} must carry the legend's size, weight, colour and tracking; got: {entry}"
        );
    }

    let mut fallback = legend_chart(LegendPosition::Bottom);
    fallback.chart_type = ChartType::Bar;
    fallback.legend_text_style = style;
    let source = chart_source(fallback);
    assert!(
        source.contains(
            "#text(size: 17pt, weight: \"bold\", fill: rgb(192, 42, 122), tracking: 1.25pt, ligatures: false, kerning: false)[Text]"
        ),
        "the unframed bar fallback uses the same legend style; got:\n{source}"
    );
}

#[test]
fn a_bottom_legend_leaves_the_plot_the_full_frame_width() {
    // The right-hand legend stole about 84pt of plot width from a chart that
    // asked for the legend underneath.
    let bottom = chart_source(legend_chart(LegendPosition::Bottom));
    let right = chart_source(legend_chart(LegendPosition::Right));

    let width_of = |source: &str| -> f64 {
        source
            .lines()
            .find(|line| line.starts_with("#box(width:"))
            .and_then(|line| {
                line.split_once("width: ")?
                    .1
                    .split_once("pt")?
                    .0
                    .parse()
                    .ok()
            })
            .expect("a plot box is emitted")
    };

    assert!(
        width_of(&bottom) < width_of(&right),
        "a bottom legend must not reserve a column beside the plot: {} vs {}",
        width_of(&bottom),
        width_of(&right)
    );
}

#[test]
fn a_left_legend_shifts_the_plot_clear_of_it() {
    // The plot must move right by what the legend reserves, or the two overlap.
    let left = chart_source(legend_chart(LegendPosition::Left));
    let entries = emitted_legend_entries(&left);
    let first_bar_x: f64 = left
        .lines()
        .find(|line| line.contains("rect(width:"))
        .and_then(|line| line.split_once("dx: ")?.1.split_once("pt")?.0.parse().ok())
        .expect("a bar is emitted");

    assert!(
        entries[0].0 < first_bar_x,
        "a left legend sits clear of the plot: legend at {}, first bar at {first_bar_x}",
        entries[0].0
    );
}

// ----- Declared series and point fills (issue #535) -----

#[test]
fn a_declared_series_fill_reaches_the_bars() {
    // The palette's first entry is rgb(68, 114, 196); the file says 4F81BD.
    let chart = Chart {
        chart_type: ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Production LOC by layer".to_string()),
        categories: vec!["parser".to_string(), "render".to_string()],
        series: vec![ChartSeries {
            name: Some("LOC".to_string()),
            values: vec![23334.0, 8331.0],
            fill: Some(Color::new(0x4f, 0x81, 0xbd)),
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    };

    let source = chart_source(chart);

    assert!(
        source.contains("rgb(79, 129, 189)"),
        "the declared 4F81BD must reach the bars, got:\n{source}"
    );
    assert!(
        !source.contains("rgb(68, 114, 196)"),
        "the palette must not override a declared fill, got:\n{source}"
    );
}

#[test]
fn a_series_without_a_fill_still_takes_the_palette() {
    // Control: the palette remains the fallback, so this is not a blanket
    // change to how charts are coloured.
    let chart = Chart {
        chart_type: ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["parser".to_string(), "render".to_string()],
        series: vec![ChartSeries {
            name: Some("LOC".to_string()),
            values: vec![23334.0, 8331.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    };

    let source = chart_source(chart);

    assert!(
        source.contains("rgb(68, 114, 196)"),
        "an undeclared series keeps the palette, got:\n{source}"
    );
}

#[test]
fn per_point_fills_colour_each_bar_separately() {
    let chart = Chart {
        chart_type: ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["DOCX".to_string(), "PPTX".to_string(), "XLSX".to_string()],
        series: vec![ChartSeries {
            name: Some("Fixtures".to_string()),
            values: vec![115.0, 92.0, 138.0],
            fill: Some(Color::new(0x11, 0x11, 0x11)),
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: vec![
                Some(Color::new(0x4f, 0x81, 0xbd)),
                Some(Color::new(0xc0, 0x50, 0x4d)),
                Some(Color::new(0x9b, 0xbb, 0x59)),
            ],
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    };

    let source = chart_source(chart);

    for expected in ["rgb(79, 129, 189)", "rgb(192, 80, 77)", "rgb(155, 187, 89)"] {
        assert!(
            source.contains(expected),
            "each point paints its own fill; {expected} missing from:\n{source}"
        );
    }
}

// ----- Axis titles (issue #552) -----

fn axis_titled_chart(category: Option<&str>, value: Option<&str>) -> Chart {
    Chart {
        chart_type: ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Production LOC by layer".to_string()),
        categories: vec!["parser".to_string(), "render".to_string()],
        series: vec![ChartSeries {
            name: Some("LOC".to_string()),
            values: vec![23334.0, 8331.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: category.map(str::to_string),
        value_axis_title: value.map(str::to_string),
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

#[test]
fn axis_titles_are_drawn() {
    let source = chart_source(axis_titled_chart(Some("계층"), Some("LOC")));

    assert!(
        source.contains("계층"),
        "the category axis title must be drawn: {source}"
    );
    assert!(
        source.contains("rotate(-90deg"),
        "the value axis title runs down the left edge: {source}"
    );
}

#[test]
fn an_untitled_axis_reserves_no_band() {
    // Control: a chart with no axis titles keeps its old geometry, so the
    // gutters are spent only when there is something to put in them.
    let titled = chart_source(axis_titled_chart(Some("계층"), Some("LOC")));
    let untitled = chart_source(axis_titled_chart(None, None));

    let box_width = |source: &str| -> f64 {
        source
            .lines()
            .find(|line| line.starts_with("#box(width:"))
            .and_then(|line| {
                line.split_once("width: ")?
                    .1
                    .split_once("pt")?
                    .0
                    .parse()
                    .ok()
            })
            .expect("a plot box is emitted")
    };

    assert!(
        box_width(&titled) > box_width(&untitled),
        "the value axis title widens the box: {} vs {}",
        box_width(&titled),
        box_width(&untitled)
    );
    assert!(
        !untitled.contains("rotate(-90deg"),
        "nothing is rotated when no axis is titled: {untitled}"
    );
}

#[test]
fn each_axis_title_is_independent() {
    let value_only = chart_source(axis_titled_chart(None, Some("LOC")));

    assert!(value_only.contains("rotate(-90deg"));
    assert!(
        !value_only.contains("계층"),
        "an untitled category axis draws nothing: {value_only}"
    );
}

// ----- Data labels (issue #547) -----

fn labelled_chart(labels: DataLabels) -> Chart {
    Chart {
        chart_type: ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["DOCX".to_string(), "PPTX".to_string()],
        series: vec![ChartSeries {
            name: Some("Text".to_string()),
            values: vec![4.0, 2.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: labels,
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Stacked,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

#[test]
fn show_val_prints_one_label_per_point() {
    let source = chart_source(labelled_chart(DataLabels {
        show_value: true,
        ..DataLabels::default()
    }));
    let labels = source.matches("weight: \"bold\", fill: white").count();

    assert_eq!(labels, 2, "one label per plotted point, got:\n{source}");
}

#[test]
fn a_series_without_dlbls_draws_no_labels() {
    // Control: the label pass is driven by the file, not switched on for all.
    let source = chart_source(labelled_chart(DataLabels::default()));

    assert!(
        !source.contains("weight: \"bold\", fill: white"),
        "no labels without dLbls, got:\n{source}"
    );
}

#[test]
fn the_enabled_parts_are_joined_by_the_separator() {
    let source = chart_source(labelled_chart(DataLabels {
        show_value: true,
        show_category: true,
        show_series: true,
        separator: "; ".to_string(),
        position: crate::ir::DataLabelPosition::Center,
        position_stated: false,
        ..DataLabels::default()
    }));

    assert!(
        source.contains("Text; DOCX; 4"),
        "series, category, then value, joined by the separator, got:\n{source}"
    );
}

#[test]
fn percent_labels_are_a_share_of_the_category() {
    let source = chart_source(labelled_chart(DataLabels {
        show_percent: true,
        ..DataLabels::default()
    }));

    // A lone series takes the whole category.
    assert!(
        source.contains("100%"),
        "the only series in a category is all of it, got:\n{source}"
    );
}

// ----- Pie geometry (issue #533) -----

fn pie_chart(values: Vec<f64>) -> Chart {
    Chart {
        chart_type: ChartType::Pie,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Fixture documents by format".to_string()),
        categories: vec!["DOCX".to_string(), "PPTX".to_string(), "XLSX".to_string()],
        series: vec![ChartSeries {
            name: None,
            values,
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

#[test]
fn a_pie_chart_draws_wedges_not_a_table() {
    let source = chart_source(pie_chart(vec![115.0, 92.0, 138.0]));

    assert_eq!(
        source.matches("curve(fill:").count(),
        3,
        "one wedge per slice, got:\n{source}"
    );
    assert!(
        !source.contains("Pie Chart"),
        "the type-label fallback is gone, got:\n{source}"
    );
}

#[test]
fn a_pie_skips_slices_with_no_value() {
    // A zero slice has no wedge to draw, but keeps its legend entry.
    let source = chart_source(pie_chart(vec![115.0, 0.0, 138.0]));

    assert_eq!(source.matches("curve(fill:").count(), 2);
    assert!(source.contains("PPTX"), "the legend still lists it");
}

#[test]
fn an_empty_pie_falls_back_to_the_table() {
    // Control: with nothing to apportion there is no pie, so the data table
    // still carries the categories.
    let source = chart_source(pie_chart(vec![0.0, 0.0, 0.0]));

    assert!(
        !source.contains("curve(fill:"),
        "no wedges without values, got:\n{source}"
    );
    assert!(source.contains("Pie Chart"), "the fallback still runs");
}

#[test]
fn the_first_wedge_starts_at_twelve_oclock() {
    // Office sweeps clockwise from the top; the first arc vertex is therefore
    // directly above the centre.
    let source = chart_source(pie_chart(vec![115.0, 92.0, 138.0]));
    let first_path: &str = source
        .lines()
        .find(|line| line.contains("curve(fill:"))
        .expect("a wedge is drawn");

    // `curve.move((cx, cy)), curve.line((cx, cy - r)), …` — the centre, then
    // the top.
    let (points, _) = curve_vertices(first_path);
    let (centre, start): ((f64, f64), (f64, f64)) = (points[0], points[1]);

    assert!(
        (start.0 - centre.0).abs() < 0.01,
        "the first vertex sits directly above the centre: {start:?} vs {centre:?}"
    );
    assert!(
        start.1 < centre.1,
        "and above it, not below: {start:?} vs {centre:?}"
    );
}

#[test]
fn wedge_colours_follow_the_declared_data_point_fills() {
    let mut chart = pie_chart(vec![115.0, 92.0, 138.0]);
    chart.series[0].point_fills = vec![
        Some(Color::new(0x4f, 0x81, 0xbd)),
        Some(Color::new(0xc0, 0x50, 0x4d)),
        Some(Color::new(0x9b, 0xbb, 0x59)),
    ];

    let source = chart_source(chart);

    for expected in ["rgb(79, 129, 189)", "rgb(192, 80, 77)", "rgb(155, 187, 89)"] {
        assert!(
            source.contains(expected),
            "wedge colour {expected} missing from:\n{source}"
        );
    }
}

// ----- Pie data labels (issue #570) -----

#[test]
fn a_pie_draws_a_label_on_each_wedge() {
    let mut chart = pie_chart(vec![115.0, 92.0, 138.0]);
    chart.series[0].data_labels = DataLabels {
        show_value: true,
        show_category: true,
        show_percent: true,
        separator: "; ".to_string(),
        position: crate::ir::DataLabelPosition::Center,
        position_stated: false,
        ..DataLabels::default()
    };

    let source = chart_source(chart);

    assert_eq!(
        source.matches("weight: \"bold\", fill: white").count(),
        3,
        "one label per wedge, got:\n{source}"
    );
    assert!(
        source.contains("DOCX; 115; 33%"),
        "category, value and share, joined by the separator, got:\n{source}"
    );
}

#[test]
fn a_pie_without_dlbls_draws_no_wedge_labels() {
    // Control: the labels are driven by the file, as on the axis plot.
    let source = chart_source(pie_chart(vec![115.0, 92.0, 138.0]));

    assert!(
        !source.contains("weight: \"bold\", fill: white"),
        "no labels without dLbls, got:\n{source}"
    );
}

#[test]
fn a_zero_slice_carries_no_label() {
    let mut chart = pie_chart(vec![115.0, 0.0, 138.0]);
    chart.series[0].data_labels = DataLabels {
        show_value: true,
        ..DataLabels::default()
    };

    let source = chart_source(chart);

    assert_eq!(
        source.matches("weight: \"bold\", fill: white").count(),
        2,
        "a slice with no wedge has nothing to label, got:\n{source}"
    );
}

/// An automatic major gridline is PowerPoint's 0.75pt `#868686`, not a lighter
/// hairline.
///
/// `c:majorGridlines` with no `c:spPr` leaves both sides drawing their own
/// default. Ours was 0.6pt `#C8C8C8`, which puts roughly a quarter of the ink
/// on each line and leaves the grid barely visible against a white plot area
/// (issue #673).
#[test]
fn test_chart_default_gridline_matches_powerpoint() {
    let doc = make_doc(vec![make_flow_page(vec![Block::Chart(Box::new(Chart {
        chart_type: ChartType::Bar,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["Q1".to_string(), "Q2".to_string()],
        series: vec![ChartSeries {
            name: Some("Revenue".to_string()),
            values: vec![100.0, 250.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }))])]);

    let source = generate_typst(&doc).unwrap().source;
    assert!(
        source.contains("stroke: 0.75pt + rgb(134, 134, 134)"),
        "gridlines should be PowerPoint's 0.75pt #868686, got:\n{source}"
    );
    assert!(
        !source.contains("rgb(200, 200, 200)"),
        "the old #C8C8C8 gridline default must be gone, got:\n{source}"
    );
    assert!(
        !source.contains("rgb(120, 120, 120)"),
        "the old #787878 axis-line default must be gone, got:\n{source}"
    );
}

// ----- Axis lines and major tick marks (issue #672) -----

/// One `line(...)` the generator placed: its top-left corner and the offset of
/// its far end, in points.
#[derive(Debug, Clone, Copy)]
struct PlacedLine {
    dx: f64,
    dy: f64,
    end_x: f64,
    end_y: f64,
}

/// Read the point measurement a slice starts with: `"12.5pt, …"` becomes 12.5.
fn leading_pt(text: &str) -> Option<f64> {
    text.split_once("pt")?.0.trim().parse::<f64>().ok()
}

/// Every plot segment the source places, in the order written.
///
/// Only the chart's chrome counts: gridlines, axis lines and tick marks all
/// take `CHART_AUTOMATIC_LINE`. A line series' legend key is also a `line`, but
/// it is drawn in the series colour at the series weight and is not plot
/// geometry, so counting it made every tick census see one tick too many
/// (#801).
///
/// The chart-area outline carries the same stroke (#637) and so passes the
/// substring filter, but it is a `box`, not a `line`, and the `line(end: (`
/// parse below drops it. Both conditions are load-bearing — neither alone
/// selects the plot segments.
fn emitted_lines(source: &str) -> Vec<PlacedLine> {
    source
        .lines()
        .filter(|line| line.contains(CHART_AUTOMATIC_LINE))
        .filter_map(|line| {
            let (placement, end) = line.split_once("line(end: (")?;
            Some(PlacedLine {
                dx: leading_pt(placement.split_once("dx: ")?.1)?,
                dy: leading_pt(placement.split_once("dy: ")?.1)?,
                end_x: leading_pt(end)?,
                end_y: leading_pt(end.split_once(", ")?.1)?,
            })
        })
        .collect()
}

/// Whether two point measurements are the same length.
fn same_length(left: f64, right: f64) -> bool {
    (left - right).abs() < 1e-6
}

/// The plotting rectangle, as `(x, y, width, height)`, read off the segments
/// the chart drew rather than off the generator's layout constants: the
/// gridlines and both axis lines each run a whole side of the plot, so the
/// longest horizontal and vertical segments give its extents and the shorter
/// tick marks fall out.
fn plot_rect(lines: &[PlacedLine]) -> (f64, f64, f64, f64) {
    let width: f64 = lines.iter().map(|line| line.end_x).fold(0.0, f64::max);
    let height: f64 = lines.iter().map(|line| line.end_y).fold(0.0, f64::max);
    let x: f64 = lines
        .iter()
        .filter(|line| same_length(line.end_x, width))
        .map(|line| line.dx)
        .fold(f64::INFINITY, f64::min);
    let y: f64 = lines
        .iter()
        .filter(|line| same_length(line.end_y, height))
        .map(|line| line.dy)
        .fold(f64::INFINITY, f64::min);
    (x, y, width, height)
}

/// The tick marks crossing the axis line under the plot and the one down its
/// left edge: every segment too short to be a gridline or an axis line.
///
/// Which axis owns which edge depends on the orientation, so the split is by
/// edge. A column chart's bottom edge is its category axis; a bar chart's is
/// its value axis.
fn tick_marks_by_edge(
    lines: &[PlacedLine],
    plot: (f64, f64, f64, f64),
) -> (Vec<PlacedLine>, Vec<PlacedLine>) {
    let (_, _, width, height) = plot;
    let under: Vec<PlacedLine> = lines
        .iter()
        .filter(|line| same_length(line.end_x, 0.0) && line.end_y < height)
        .copied()
        .collect();
    let beside: Vec<PlacedLine> = lines
        .iter()
        .filter(|line| same_length(line.end_y, 0.0) && line.end_x < width)
        .copied()
        .collect();
    (under, beside)
}

/// The categories `tick_mark_chart` plots, so a test can look their labels up.
const TICK_MARK_CATEGORIES: [&str; 3] = ["Mon", "Tue", "Wed"];

/// A three-category chart carrying the tick marks each axis asks for.
fn tick_mark_chart(
    chart_type: ChartType,
    category_axis_major_tick_mark: AxisTickMark,
    value_axis_major_tick_mark: AxisTickMark,
) -> Chart {
    Chart {
        chart_type,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Weekly Throughput".to_string()),
        categories: TICK_MARK_CATEGORIES.map(str::to_string).to_vec(),
        series: vec![ChartSeries {
            name: Some("Builds".to_string()),
            values: vec![4.0, 8.0, 6.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark,
        value_axis_major_tick_mark,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

/// One placed `box(...)`: its top-left corner and the extent it was given, in
/// points. A box the generator left unsized vertically reports zero height.
#[derive(Debug, Clone, Copy)]
struct PlacedBox {
    dx: f64,
    dy: f64,
    width: f64,
    height: f64,
}

/// Where the generator placed the box printing `text`.
fn placed_box_holding(source: &str, text: &str) -> PlacedBox {
    let needle: String = format!("[{text}]");
    let line: &str = source
        .lines()
        .find(|line| line.contains("box(width: ") && line.contains(&needle))
        .unwrap_or_else(|| panic!("nothing prints {text} in:\n{source}"));
    let read = |prefix: &str| -> Option<f64> { leading_pt(line.split_once(prefix)?.1) };
    PlacedBox {
        dx: read("dx: ").expect("a placed box carries a dx"),
        dy: read("dy: ").expect("a placed box carries a dy"),
        width: read("box(width: ").expect("a placed box carries a width"),
        height: read("height: ").unwrap_or(0.0),
    }
}

/// The offsets in ascending order, with the coincident ones folded together —
/// the zero gridline and the axis line beside it are drawn as two segments over
/// the same offset.
fn sorted_unique(offsets: impl IntoIterator<Item = f64>) -> Vec<f64> {
    let mut sorted: Vec<f64> = offsets.into_iter().collect();
    sorted.sort_by(f64::total_cmp);
    sorted.dedup_by(|left, right| same_length(*left, *right));
    sorted
}

/// Where along the value axis each gridline runs, one per major unit.
///
/// A column or line chart draws them horizontally across the whole plot, a bar
/// chart vertically down its whole height.
fn gridline_offsets(
    lines: &[PlacedLine],
    plot: (f64, f64, f64, f64),
    value_axis_runs_under_the_plot: bool,
) -> Vec<f64> {
    let (_, _, width, height) = plot;
    sorted_unique(lines.iter().filter_map(|line| {
        if value_axis_runs_under_the_plot {
            (same_length(line.end_x, 0.0) && same_length(line.end_y, height)).then_some(line.dx)
        } else {
            (same_length(line.end_y, 0.0) && same_length(line.end_x, width)).then_some(line.dy)
        }
    }))
}

/// Where along its axis each tick sits. A tick runs across its axis, so the
/// coordinate that stays put over the tick's own length is the one that places
/// it along the axis — whichever side of the line the tick reaches from.
fn tick_offsets(ticks: &[PlacedLine], along_the_bottom_edge: bool) -> Vec<f64> {
    sorted_unique(ticks.iter().map(|tick| {
        if along_the_bottom_edge {
            tick.dx
        } else {
            tick.dy
        }
    }))
}

/// A value tick marks the same major unit its gridline does, so the two land on
/// the same offset along the axis. Counting ticks alone would pass an
/// implementation that drew the right number of them in the wrong places.
fn assert_value_ticks_sit_on_their_gridlines(source: &str, value_axis_runs_under_the_plot: bool) {
    let lines: Vec<PlacedLine> = emitted_lines(source);
    let plot: (f64, f64, f64, f64) = plot_rect(&lines);
    let (under, beside) = tick_marks_by_edge(&lines, plot);
    let value_ticks: &[PlacedLine] = if value_axis_runs_under_the_plot {
        &under
    } else {
        &beside
    };

    let gridlines: Vec<f64> = gridline_offsets(&lines, plot, value_axis_runs_under_the_plot);
    let ticks: Vec<f64> = tick_offsets(value_ticks, value_axis_runs_under_the_plot);

    assert!(!gridlines.is_empty(), "no gridlines drawn in:\n{source}");
    assert_eq!(
        ticks.len(),
        gridlines.len(),
        "one value tick per gridline; ticks at {ticks:?} against gridlines at {gridlines:?}\n{source}"
    );
    for (tick, gridline) in ticks.iter().zip(&gridlines) {
        assert!(
            same_length(*tick, *gridline),
            "a value tick at {tick} misses its gridline at {gridline}; ticks {ticks:?} against gridlines {gridlines:?}\n{source}"
        );
    }
}

/// The category ticks are the boundaries of the bands the labels sit in the
/// middle of: evenly spaced along the axis, with every label's centre exactly
/// midway between two neighbouring ticks.
///
/// This is what `<c:crossBetween val="between"/>` means, and it is what pins the
/// ticks to the layout: ticks placed by a rule of their own can still come out
/// evenly spaced and correctly counted while sitting nowhere near a label.
fn assert_category_ticks_bound_the_labels(
    source: &str,
    categories: &[&str],
    value_axis_runs_under_the_plot: bool,
) {
    let lines: Vec<PlacedLine> = emitted_lines(source);
    let plot: (f64, f64, f64, f64) = plot_rect(&lines);
    let (under, beside) = tick_marks_by_edge(&lines, plot);
    // A bar chart's categories run down the left edge; every other orientation
    // lays them along the bottom.
    let category_axis_is_the_bottom_edge: bool = !value_axis_runs_under_the_plot;
    let category_ticks: &[PlacedLine] = if category_axis_is_the_bottom_edge {
        &under
    } else {
        &beside
    };

    let boundaries: Vec<f64> = tick_offsets(category_ticks, category_axis_is_the_bottom_edge);
    assert_eq!(
        boundaries.len(),
        categories.len() + 1,
        "one tick per band boundary, so one more than the categories; got {boundaries:?}\n{source}"
    );
    let pitch: f64 = boundaries[1] - boundaries[0];
    for pair in boundaries.windows(2) {
        assert!(
            same_length(pair[1] - pair[0], pitch),
            "the bands the ticks bound must all be the same width; got {boundaries:?}\n{source}"
        );
    }

    let band_centres: Vec<f64> = boundaries
        .windows(2)
        .map(|pair| (pair[0] + pair[1]) / 2.0)
        .collect();
    let label_centres: Vec<f64> = sorted_unique(categories.iter().map(|category| {
        let label: PlacedBox = placed_box_holding(source, category);
        if category_axis_is_the_bottom_edge {
            label.dx + label.width / 2.0
        } else {
            label.dy + label.height / 2.0
        }
    }));
    assert_eq!(label_centres.len(), band_centres.len());
    for (label, band) in label_centres.iter().zip(&band_centres) {
        assert!(
            same_length(*label, *band),
            "a category label centred on {label} is not in the middle of a band; labels {label_centres:?} against bands bounded by {boundaries:?}\n{source}"
        );
    }
}

/// Both sides of the plot carry an axis line. The value axis was never stroked
/// for a bar or a column chart, whichever edge it owned (issue #672).
fn assert_both_axis_lines(source: &str) {
    let lines: Vec<PlacedLine> = emitted_lines(source);
    let (plot_x, plot_y, plot_w, plot_h) = plot_rect(&lines);

    assert!(
        lines.iter().any(|line| same_length(line.dx, plot_x)
            && same_length(line.dy, plot_y)
            && same_length(line.end_x, 0.0)
            && same_length(line.end_y, plot_h)),
        "no axis line down the plot's left edge at x={plot_x}, y={plot_y}..{}; got:\n{source}",
        plot_y + plot_h
    );
    assert!(
        lines.iter().any(|line| same_length(line.dx, plot_x)
            && same_length(line.dy, plot_y + plot_h)
            && same_length(line.end_x, plot_w)
            && same_length(line.end_y, 0.0)),
        "no axis line along the plot's bottom edge at y={}, x={plot_x}..{}; got:\n{source}",
        plot_y + plot_h,
        plot_x + plot_w
    );
}

#[test]
fn a_column_chart_strokes_both_of_its_axis_lines() {
    assert_both_axis_lines(&chart_source(tick_mark_chart(
        ChartType::Column,
        AxisTickMark::Outside,
        AxisTickMark::Outside,
    )));
}

#[test]
fn a_horizontal_bar_chart_strokes_both_of_its_axis_lines() {
    // Triangulation: the orientation swaps which axis owns which edge, so one
    // hardcoded edge cannot satisfy both charts.
    assert_both_axis_lines(&chart_source(tick_mark_chart(
        ChartType::Bar,
        AxisTickMark::Outside,
        AxisTickMark::Outside,
    )));
}

#[test]
fn a_line_chart_strokes_both_of_its_axis_lines() {
    assert_both_axis_lines(&chart_source(tick_mark_chart(
        ChartType::Line,
        AxisTickMark::Outside,
        AxisTickMark::Outside,
    )));
}

/// Each axis ticks every major unit, and the category axis ticks every band
/// boundary — `<c:crossBetween val="between"/>` gives three categories four of
/// them, as Excel and PowerPoint both draw.
fn assert_tick_counts(source: &str, value_axis_runs_under_the_plot: bool) {
    let lines: Vec<PlacedLine> = emitted_lines(source);
    let plot: (f64, f64, f64, f64) = plot_rect(&lines);
    let (under, beside) = tick_marks_by_edge(&lines, plot);

    let category_boundaries: usize = TICK_MARK_CATEGORIES.len() + 1;
    let major_units: usize = emitted_axis_ticks(source).len();
    assert_eq!(major_units, 10, "values 4/8/6 scale to ticks 0..9 by 1");
    let (expected_under, expected_beside) = if value_axis_runs_under_the_plot {
        (major_units, category_boundaries)
    } else {
        (category_boundaries, major_units)
    };

    assert_eq!(
        under.len(),
        expected_under,
        "tick marks under the plot: {under:#?}\n{source}"
    );
    assert_eq!(
        beside.len(),
        expected_beside,
        "tick marks left of the plot: {beside:#?}\n{source}"
    );
}

#[test]
fn a_column_chart_ticks_every_major_unit_and_every_category_boundary() {
    // A column chart's value axis runs down the left edge, so its major-unit
    // ticks are the ones beside the plot.
    assert_tick_counts(
        &chart_source(tick_mark_chart(
            ChartType::Column,
            AxisTickMark::Outside,
            AxisTickMark::Outside,
        )),
        false,
    );
}

#[test]
fn a_horizontal_bar_chart_ticks_the_edges_the_other_way_round() {
    assert_tick_counts(
        &chart_source(tick_mark_chart(
            ChartType::Bar,
            AxisTickMark::Outside,
            AxisTickMark::Outside,
        )),
        true,
    );
}

#[test]
fn a_line_chart_ticks_both_of_its_axes() {
    assert_tick_counts(
        &chart_source(tick_mark_chart(
            ChartType::Line,
            AxisTickMark::Outside,
            AxisTickMark::Outside,
        )),
        false,
    );
}

/// A chart's ticks land on the geometry the same chart drew, on both axes.
fn assert_ticks_match_the_plot(chart_type: ChartType, value_axis_runs_under_the_plot: bool) {
    let source: String = chart_source(tick_mark_chart(
        chart_type,
        AxisTickMark::Outside,
        AxisTickMark::Outside,
    ));
    assert_value_ticks_sit_on_their_gridlines(&source, value_axis_runs_under_the_plot);
    assert_category_ticks_bound_the_labels(
        &source,
        &TICK_MARK_CATEGORIES,
        value_axis_runs_under_the_plot,
    );
}

#[test]
fn a_column_chart_puts_every_tick_on_the_geometry_it_marks() {
    assert_ticks_match_the_plot(ChartType::Column, false);
}

#[test]
fn a_horizontal_bar_chart_puts_every_tick_on_the_geometry_it_marks() {
    assert_ticks_match_the_plot(ChartType::Bar, true);
}

/// The `dx:` an emitted horizontal category-label `#place` carries for the
/// given category text, or `None` if no such line was written.
fn category_label_box_dx(source: &str, category: &str) -> Option<f64> {
    let suffix: String = format!("[{category}]]])");
    let line: &str = source
        .lines()
        .find(|line| line.contains("align(right + horizon)") && line.ends_with(&suffix))?;
    let after_dx: &str = line.split_once("dx: ")?.1;
    after_dx.split_once("pt")?.0.parse::<f64>().ok()
}

#[test]
fn an_excel_bar_category_label_sits_the_measured_offset_right_of_powerpoints() {
    // Native Microsoft Excel 16.112 grid on
    // `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx` (#1620), printed at
    // this sheet's own 0.78 fit scale: patching the bar category axis size to
    // 8, 10, 14 and 18pt and comparing a fresh native export against the
    // PowerPoint-calibrated box shared by `chart_category_label_box_w` gives
    // a label origin 0.2083, 0.2610, 0.3677 and 0.4733 **printed** points
    // right of that box at each size. `framed_chart_source` below calls
    // `generate_chart_in` directly, bypassing the fitted sheet's
    // `#scale(print_scale)` wrapper `write_placed_sheet_anchor` applies at
    // runtime, so `excel_dx - pptx_dx` here is a sheet-space (unscaled)
    // delta; each literal is the printed measurement divided by this sheet's
    // own 0.78 print scale to compare like with like. The two extreme sizes
    // probed pin the measured rows themselves (not a recomputation of the
    // constant under test) so a wrong slope or intercept in
    // `EXCEL_BAR_CATEGORY_LABEL_X_SHIFT_EM` fails this test — see
    // `assets/validation/issue-1620/native-category-label-measurements.json`.
    let print_scale: f64 = 0.78;
    for (size_pt, measured_native_printed_dx0_pt) in [(8.0, 0.2083), (18.0, 0.4733)] {
        let measured_native_sheet_dx0_pt: f64 = measured_native_printed_dx0_pt / print_scale;
        let mut chart = two_series_bar_chart(Vec::new());
        chart.categories = vec!["financial aid".to_string(), "other".to_string()];
        chart.series[0].values = vec![10.0, 2.0];
        chart.series.truncate(1);
        chart.category_axis_text_style.size_pt = Some(size_pt);
        chart.text_style.size_pt = Some(size_pt);

        chart.host = crate::ir::ChartHost::Presentation;
        let pptx_source: String = framed_chart_source(&chart, 400.0, 250.0);
        chart.host = crate::ir::ChartHost::Spreadsheet;
        let excel_source: String = framed_chart_source(&chart, 400.0, 250.0);

        let pptx_dx: f64 = category_label_box_dx(&pptx_source, "financial aid")
            .expect("PowerPoint host emits the category label");
        let excel_dx: f64 = category_label_box_dx(&excel_source, "financial aid")
            .expect("Excel host emits the category label");

        assert!(
            (excel_dx - pptx_dx - measured_native_sheet_dx0_pt).abs() < 0.01,
            "at {size_pt}pt the Excel box must sit the native-measured \
             {measured_native_sheet_dx0_pt}pt (sheet-space) right of \
             PowerPoint's shared box; got excel={excel_dx} pptx={pptx_dx}"
        );
    }
}

#[test]
fn an_excel_bar_category_label_with_no_declared_size_keeps_powerpoints_box() {
    // No native export establishes this shift for implicit (file-default)
    // text, so `excel_bar_category_label_x_shift_pt` deliberately stays at
    // 0.0 when neither the category axis nor the chart space declares a
    // size, mirroring `excel_category_label_y_shift_pt`'s same guard.
    let mut chart = two_series_bar_chart(Vec::new());
    chart.categories = vec!["financial aid".to_string(), "other".to_string()];
    chart.series[0].values = vec![10.0, 2.0];
    chart.series.truncate(1);
    assert!(chart.category_axis_text_style.size_pt.is_none());
    assert!(chart.text_style.size_pt.is_none());

    chart.host = crate::ir::ChartHost::Presentation;
    let pptx_source: String = framed_chart_source(&chart, 400.0, 250.0);
    chart.host = crate::ir::ChartHost::Spreadsheet;
    let excel_source: String = framed_chart_source(&chart, 400.0, 250.0);

    let pptx_dx: f64 = category_label_box_dx(&pptx_source, "financial aid")
        .expect("PowerPoint host emits the category label");
    let excel_dx: f64 = category_label_box_dx(&excel_source, "financial aid")
        .expect("Excel host emits the category label");

    assert!(
        (excel_dx - pptx_dx).abs() < 1e-9,
        "an undeclared size must keep the shared PowerPoint box; \
         got excel={excel_dx} pptx={pptx_dx}"
    );
}

#[test]
fn a_line_chart_puts_every_tick_on_the_geometry_it_marks() {
    // The line plot lays its categories out in bands of its own, so its ticks
    // have to be read off that layout rather than borrowed from the bar family.
    assert_ticks_match_the_plot(ChartType::Line, false);
}

#[test]
fn an_axis_asking_for_no_tick_marks_gets_none() {
    // Triangulation against drawing ticks unconditionally, and against reading
    // one axis' setting for both: only the category axis goes quiet here.
    let source: String = chart_source(tick_mark_chart(
        ChartType::Column,
        AxisTickMark::None,
        AxisTickMark::Outside,
    ));
    let lines: Vec<PlacedLine> = emitted_lines(&source);
    let plot: (f64, f64, f64, f64) = plot_rect(&lines);
    let (under, beside) = tick_marks_by_edge(&lines, plot);

    assert!(
        under.is_empty(),
        "a category axis asking for no tick marks must draw none, got {under:#?}\n{source}"
    );
    assert!(
        !beside.is_empty(),
        "the value axis still asked for tick marks, got:\n{source}"
    );
}

#[test]
fn inward_tick_marks_reach_into_the_plot_and_crossing_ones_both_ways() {
    // `in` and `out` mirror each other about the axis line and `cross` is
    // both, so the mode has to steer the geometry rather than only decide
    // whether a segment is drawn at all.
    let left_edge_ticks = |mark: AxisTickMark| -> (f64, Vec<PlacedLine>) {
        let source: String =
            chart_source(tick_mark_chart(ChartType::Column, AxisTickMark::None, mark));
        let lines: Vec<PlacedLine> = emitted_lines(&source);
        let plot: (f64, f64, f64, f64) = plot_rect(&lines);
        (plot.0, tick_marks_by_edge(&lines, plot).1)
    };

    let (axis_x, outward) = left_edge_ticks(AxisTickMark::Outside);
    let (_, inward) = left_edge_ticks(AxisTickMark::Inside);
    let (_, crossing) = left_edge_ticks(AxisTickMark::Cross);

    assert!(!outward.is_empty() && !inward.is_empty() && !crossing.is_empty());
    assert!(
        outward
            .iter()
            .all(|tick| tick.dx < axis_x && same_length(tick.dx + tick.end_x, axis_x)),
        "an outward tick ends on the axis line at x={axis_x}, got {outward:#?}"
    );
    assert!(
        inward
            .iter()
            .all(|tick| same_length(tick.dx, axis_x) && tick.end_x > 0.0),
        "an inward tick starts on the axis line at x={axis_x}, got {inward:#?}"
    );
    assert!(
        crossing
            .iter()
            .all(|tick| tick.dx < axis_x && tick.dx + tick.end_x > axis_x),
        "a crossing tick straddles the axis line at x={axis_x}, got {crossing:#?}"
    );
    assert_eq!(
        outward.len(),
        crossing.len(),
        "every mode ticks the same major units"
    );
    assert!(
        crossing[0].end_x > outward[0].end_x,
        "a crossing tick is longer than a one-sided one: {crossing:#?} vs {outward:#?}"
    );
}

/// A column chart with one of its axes switched off by `<c:delete val="1"/>`,
/// both still asking for outward ticks — which is what Office leaves behind
/// when a user unticks an axis rather than setting its tick marks to `none`.
fn chart_with_deleted_axis(category_deleted: bool, value_deleted: bool) -> Chart {
    let mut chart: Chart = tick_mark_chart(
        ChartType::Column,
        AxisTickMark::Outside,
        AxisTickMark::Outside,
    );
    chart.category_axis_deleted = category_deleted;
    chart.value_axis_deleted = value_deleted;
    chart
}

#[test]
fn a_deleted_value_axis_draws_no_line_no_ticks_and_no_labels() {
    let drawn: String = chart_source(chart_with_deleted_axis(false, false));
    let hidden: String = chart_source(chart_with_deleted_axis(false, true));
    // The gutters do not move when an axis goes, so the plot the deleted chart
    // draws into is the one the drawn chart reports.
    let plot: (f64, f64, f64, f64) = plot_rect(&emitted_lines(&drawn));
    let (plot_x, plot_y, _, plot_h) = plot;
    let lines: Vec<PlacedLine> = emitted_lines(&hidden);
    let (under, beside) = tick_marks_by_edge(&lines, plot);

    assert!(
        !lines.iter().any(|line| same_length(line.dx, plot_x)
            && same_length(line.dy, plot_y)
            && same_length(line.end_y, plot_h)),
        "a deleted value axis must not stroke the left edge it owns; got:\n{hidden}"
    );
    assert!(
        beside.is_empty(),
        "a deleted value axis must not tick, whatever `<c:majorTickMark>` still says; got {beside:#?}\n{hidden}"
    );
    assert!(
        emitted_axis_ticks(&hidden).is_empty(),
        "a deleted value axis must not label its units; got:\n{hidden}"
    );
    // Gridlines are a chart element of their own — deleting the axis leaves
    // them standing — and the category axis is untouched.
    assert_eq!(
        gridline_offsets(&lines, plot, false),
        gridline_offsets(&emitted_lines(&drawn), plot, false),
        "the gridlines belong to the chart, not to the axis switched off"
    );
    assert_eq!(
        under.len(),
        TICK_MARK_CATEGORIES.len() + 1,
        "the category axis still ticks every band boundary; got {under:#?}\n{hidden}"
    );
}

#[test]
fn a_deleted_category_axis_takes_only_its_own_furniture_with_it() {
    // Triangulation against one flag standing for both axes, and against the
    // deletion reaching further than the axis it names.
    let drawn: String = chart_source(chart_with_deleted_axis(false, false));
    let hidden: String = chart_source(chart_with_deleted_axis(true, false));
    let plot: (f64, f64, f64, f64) = plot_rect(&emitted_lines(&drawn));
    let (plot_x, plot_y, plot_w, plot_h) = plot;
    let lines: Vec<PlacedLine> = emitted_lines(&hidden);
    let (under, beside) = tick_marks_by_edge(&lines, plot);

    // The zero gridline runs along the bottom edge too, so the axis line there
    // is one of two coincident segments rather than the only one.
    let bottom_edge_strokes = |source: &str| -> usize {
        emitted_lines(source)
            .iter()
            .filter(|line| same_length(line.dy, plot_y + plot_h) && same_length(line.end_x, plot_w))
            .count()
    };
    assert_eq!(
        bottom_edge_strokes(&hidden),
        bottom_edge_strokes(&drawn) - 1,
        "a deleted category axis must stop stroking the bottom edge it owns; got:\n{hidden}"
    );
    assert!(
        under.is_empty(),
        "a deleted category axis must not tick; got {under:#?}\n{hidden}"
    );
    for category in TICK_MARK_CATEGORIES {
        assert!(
            !hidden.contains(&format!("[{category}]")),
            "a deleted category axis must not label its bands, found {category} in:\n{hidden}"
        );
    }
    assert!(
        lines.iter().any(|line| same_length(line.dx, plot_x)
            && same_length(line.dy, plot_y)
            && same_length(line.end_y, plot_h)),
        "the value axis is still drawn; got:\n{hidden}"
    );
    assert_eq!(
        beside.len(),
        emitted_axis_ticks(&hidden).len(),
        "the value axis still ticks every unit it labels; got {beside:#?}\n{hidden}"
    );
}

// ----- Bar thickness from c:gapWidth and c:overlap (issue #671) -----

/// One `rect(...)` the generator placed: its top-left corner and the extent it
/// was given, in points.
#[derive(Debug, Clone, Copy)]
struct PlacedRect {
    dx: f64,
    dy: f64,
    width: f64,
    height: f64,
}

/// Every rectangle the source places, in the order written. A bar or column
/// chart draws nothing else as a rectangle, so these are exactly its bars —
/// unless a combo plot area lays a line over them whose series index takes the
/// square marker of the shape cycle, which draws one too (issue #1067).
fn emitted_rects(source: &str) -> Vec<PlacedRect> {
    source
        .lines()
        .filter_map(|line| {
            let (placement, extent) = line.split_once("rect(width: ")?;
            Some(PlacedRect {
                dx: leading_pt(placement.split_once("dx: ")?.1)?,
                dy: leading_pt(placement.split_once("dy: ")?.1)?,
                width: leading_pt(extent)?,
                height: leading_pt(extent.split_once("height: ")?.1)?,
            })
        })
        .collect()
}

/// Each bar as `(start, thickness)` along the category axis — the horizontal
/// axis for a column chart, the vertical one for a horizontal bar chart.
///
/// The generator writes one category at a time, every series within it, so the
/// first `series_count` entries share the first band.
fn bars_across_the_categories(source: &str, horizontal: bool) -> Vec<(f64, f64)> {
    emitted_rects(source)
        .into_iter()
        .map(|rect| {
            if horizontal {
                (rect.dy, rect.height)
            } else {
                (rect.dx, rect.width)
            }
        })
        .collect()
}

/// Where the first category's band starts along the category axis, read off the
/// plotting rectangle the gridlines and axis lines describe.
///
/// A column chart lays its categories out left to right from the plot's left
/// edge; a horizontal bar chart stacks them bottom-up, so its first band is the
/// last one down the plot.
fn first_band_start(source: &str, horizontal: bool, categories: usize) -> f64 {
    let (plot_x, plot_y, _, plot_h) = plot_rect(&emitted_lines(source));
    if horizontal {
        plot_y + plot_h - plot_h / categories as f64
    } else {
        plot_x
    }
}

/// The three categories every band-layout test plots, with a value per series.
const BAND_SERIES_VALUES: [[f64; 3]; 4] = [
    [4.0, 2.0, 2.0],
    [1.0, 3.0, 1.0],
    [2.0, 4.0, 3.0],
    [2.0, 2.0, 3.0],
];

/// A chart of `series_count` series over three categories, declaring `layout`.
fn band_layout_chart(
    chart_type: ChartType,
    grouping: ChartGrouping,
    series_count: usize,
    layout: BarBandLayout,
) -> Chart {
    Chart {
        chart_type,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: Some("Weekly Throughput".to_string()),
        categories: vec!["Mon".to_string(), "Tue".to_string(), "Wed".to_string()],
        series: BAND_SERIES_VALUES
            .iter()
            .take(series_count)
            .enumerate()
            .map(|(index, values)| ChartSeries {
                name: Some(format!("Line {index}")),
                values: values.to_vec(),
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            })
            .collect(),
        grouping,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: layout,
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

/// A single-series chart's band pitch and bar thickness along the category
/// axis, in points.
fn pitch_and_thickness(source: &str, horizontal: bool) -> (f64, f64) {
    let bars: Vec<(f64, f64)> = bars_across_the_categories(source, horizontal);
    assert_eq!(bars.len(), 3, "one bar per category expected, got {bars:?}");
    let pitch: f64 = (bars[1].0 - bars[0].0).abs();
    assert!(
        same_length(pitch, (bars[2].0 - bars[1].0).abs()),
        "the categories must keep an even pitch, got {bars:?}"
    );
    (pitch, bars[0].1)
}

#[test]
fn a_single_series_bar_leaves_the_gutter_its_gap_width_asks_for() {
    // `<c:gapWidth>` measures the gutter between neighbouring categories in
    // units of ONE bar, so the band holds the bar plus that fraction of it.
    // Rewriting the element in `tests/fixtures/pptx/bar-chart.pptx` and tracing
    // PowerPoint 16.0's own export put every bar within one 1/1200in device
    // quantum of band / (1 + gapWidth/100), over the whole 0..500 range, while
    // the band itself never moved.
    for gap_width_percent in [0.0, 20.0, 50.0, 100.0, 150.0, 300.0, 500.0] {
        let source: String = chart_source(band_layout_chart(
            ChartType::Column,
            ChartGrouping::Clustered,
            1,
            BarBandLayout {
                gap_width_percent,
                overlap_percent: 0.0,
            },
        ));

        let (pitch, thickness) = pitch_and_thickness(&source, false);
        let expected: f64 = pitch / (1.0 + gap_width_percent / 100.0);
        assert!(
            same_length(thickness, expected),
            "gapWidth {gap_width_percent} wants a {expected}pt bar in a {pitch}pt band, got {thickness}pt"
        );
    }
}

#[test]
fn a_horizontal_bar_chart_sizes_its_bars_the_same_way() {
    // The gap is a property of the category axis, not of the page, so turning
    // the chart on its side must not change the ratio.
    for gap_width_percent in [0.0, 90.0, 219.0, 500.0] {
        let source: String = chart_source(band_layout_chart(
            ChartType::Bar,
            ChartGrouping::Clustered,
            1,
            BarBandLayout {
                gap_width_percent,
                overlap_percent: 0.0,
            },
        ));

        let (pitch, thickness) = pitch_and_thickness(&source, true);
        let expected: f64 = pitch / (1.0 + gap_width_percent / 100.0);
        assert!(
            same_length(thickness, expected),
            "gapWidth {gap_width_percent} wants a {expected}pt bar in a {pitch}pt band, got {thickness}pt"
        );
    }
}

#[test]
fn a_chart_declaring_no_gap_width_draws_the_office_default() {
    // Excel 16.0 renders `tests/fixtures/xlsx/chart_sheet.xlsx`, which declares
    // neither element, at gapWidth 150 — so an absent declaration has to reach
    // the bars as 150, leaving each bar 1/2.5 of its band.
    let source: String = chart_source(band_layout_chart(
        ChartType::Column,
        ChartGrouping::Clustered,
        1,
        BarBandLayout::default(),
    ));

    let (pitch, thickness) = pitch_and_thickness(&source, false);
    assert!(
        same_length(thickness, pitch / 2.5),
        "the default gap leaves a {}pt bar in a {pitch}pt band, got {thickness}pt",
        pitch / 2.5
    );
}

#[test]
fn every_bar_sits_centred_in_the_band_its_category_owns() {
    // PowerPoint splits the gutter evenly on both sides of the bar rather than
    // pushing it against one edge: on `tests/fixtures/pptx/bar-chart.pptx` the
    // traced bar centres sat within 0.02pt of their band centres.
    for (chart_type, horizontal) in [(ChartType::Column, false), (ChartType::Bar, true)] {
        let source: String = chart_source(band_layout_chart(
            chart_type.clone(),
            ChartGrouping::Clustered,
            1,
            BarBandLayout {
                gap_width_percent: 100.0,
                overlap_percent: 0.0,
            },
        ));

        let (pitch, thickness) = pitch_and_thickness(&source, horizontal);
        let bars: Vec<(f64, f64)> = bars_across_the_categories(&source, horizontal);
        let lead: f64 = bars[0].0 - first_band_start(&source, horizontal, 3);
        assert!(
            same_length(lead, (pitch - thickness) / 2.0),
            "{chart_type:?} must centre its bar: a {thickness}pt bar in a {pitch}pt band wants a {}pt lead, got {lead}pt",
            (pitch - thickness) / 2.0
        );
    }
}

#[test]
fn clustered_series_slide_over_each_other_by_the_declared_overlap() {
    // `<c:overlap>` moves each series' bar a fraction of a bar over the one
    // before it, so N series need N - (N-1)*overlap bars of room plus the gap.
    // Excel 16.0 draws `tests/fixtures/xlsx/any_sheets.xlsx` (219 / -27, two
    // series) as 52.5pt bars stepping 66.7pt in a 234pt band: 234/4.46 and
    // 52.47*1.27. Sweeping the sign of the overlap across five shapes leaves no
    // single ratio that could pass.
    for (gap_width_percent, overlap_percent, series_count) in [
        (219.0, -27.0, 2),
        (150.0, 0.0, 2),
        (100.0, 50.0, 2),
        (219.0, -27.0, 3),
        (90.0, 100.0, 4),
    ] {
        let source: String = chart_source(band_layout_chart(
            ChartType::Column,
            ChartGrouping::Clustered,
            series_count,
            BarBandLayout {
                gap_width_percent,
                overlap_percent,
            },
        ));

        let bars: Vec<(f64, f64)> = bars_across_the_categories(&source, false);
        assert_eq!(
            bars.len(),
            3 * series_count,
            "one bar per series per category"
        );
        let pitch: f64 = bars[series_count].0 - bars[0].0;
        let bars_wide: f64 = series_count as f64;
        let expected: f64 = pitch
            / (bars_wide - (bars_wide - 1.0) * overlap_percent / 100.0 + gap_width_percent / 100.0);
        assert!(
            same_length(bars[0].1, expected),
            "{series_count} series at {gap_width_percent}/{overlap_percent} want a {expected}pt bar in a {pitch}pt band, got {}pt",
            bars[0].1
        );

        let step: f64 = bars[1].0 - bars[0].0;
        let expected_step: f64 = expected * (1.0 - overlap_percent / 100.0);
        assert!(
            same_length(step, expected_step),
            "an overlap of {overlap_percent} steps {expected_step}pt from one series to the next, got {step}pt"
        );

        let cluster: f64 = expected + (bars_wide - 1.0) * expected_step;
        let lead: f64 = bars[0].0 - first_band_start(&source, false, 3);
        assert!(
            same_length(lead, (pitch - cluster) / 2.0),
            "the {cluster}pt cluster sits centred in its {pitch}pt band, got a lead of {lead}pt"
        );
    }
}

#[test]
fn a_stacked_category_divides_its_band_by_the_same_law_a_clustered_one_does() {
    // Stacking does not fuse the segments into one bar: `<c:overlap>` still says
    // how far each slides over the one before it. Rewriting the element on the
    // introduction deck's four-series stacked chart (gapWidth 90) and tracing
    // PowerPoint 16.0's export gave, on a 167.6pt pitch, one 88.2pt column at
    // overlap 100 (167.64/1.9) but four 34.2pt segments stepping 34.2pt at
    // overlap 0 (167.52/4.9) — a staircase, each segment still stacked on the
    // running total. Overlap 50 gave 49.3pt stepping 24.7pt and -25 gave 29.6pt
    // stepping 37.1pt. Deleting `<c:overlap>` drew the overlap-0 geometry
    // exactly, so an absent element is 0, not the 100 Office writes beside its
    // own stacked charts.
    for grouping in [ChartGrouping::Stacked, ChartGrouping::PercentStacked] {
        for overlap_percent in [100.0, 50.0, 0.0, -25.0] {
            for gap_width_percent in [90.0, 300.0] {
                let source: String = chart_source(band_layout_chart(
                    ChartType::Column,
                    grouping,
                    4,
                    BarBandLayout {
                        gap_width_percent,
                        overlap_percent,
                    },
                ));

                let bars: Vec<(f64, f64)> = bars_across_the_categories(&source, false);
                assert_eq!(bars.len(), 12, "four segments over three categories");
                let pitch: f64 = bars[4].0 - bars[0].0;
                let overlap: f64 = overlap_percent / 100.0;
                let expected: f64 = pitch / (4.0 - 3.0 * overlap + gap_width_percent / 100.0);
                let expected_step: f64 = expected * (1.0 - overlap);
                for (index, segment) in bars[..4].iter().enumerate() {
                    assert!(
                        same_length(segment.1, expected)
                            && same_length(segment.0, bars[0].0 + index as f64 * expected_step),
                        "{grouping:?} at {gap_width_percent}/{overlap_percent} wants {expected}pt segments stepping {expected_step}pt, got {bars:?}"
                    );
                }

                let cluster: f64 = expected + 3.0 * expected_step;
                let lead: f64 = bars[0].0 - first_band_start(&source, false, 3);
                assert!(
                    same_length(lead, (pitch - cluster) / 2.0),
                    "the {cluster}pt stack sits centred in its {pitch}pt band, got a lead of {lead}pt"
                );
            }
        }
    }
}

/// The Office 2007 accents both audited fixtures declare (issue #670).
fn office_2007_accents() -> Vec<crate::ir::Color> {
    vec![
        crate::ir::Color::new(0x4F, 0x81, 0xBD),
        crate::ir::Color::new(0xC0, 0x50, 0x4D),
        crate::ir::Color::new(0x9B, 0xBB, 0x59),
        crate::ir::Color::new(0x80, 0x64, 0xA2),
        crate::ir::Color::new(0x4B, 0xAC, 0xC6),
        crate::ir::Color::new(0xF7, 0x96, 0x46),
    ]
}

fn two_series_bar_chart(theme_accent_colors: Vec<crate::ir::Color>) -> Chart {
    Chart {
        chart_type: ChartType::Bar,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["Q1".to_string()],
        series: vec![
            ChartSeries {
                name: Some("Revenue".to_string()),
                values: vec![100.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
            ChartSeries {
                name: Some("Cost".to_string()),
                values: vec![60.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
        ],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors,
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

#[test]
fn test_automatic_series_colors_come_from_the_file_theme() {
    let source = chart_source(two_series_bar_chart(office_2007_accents()));

    assert!(
        source.contains("rgb(79, 129, 189)"),
        "series 1 must take the theme's accent1, got:\n{source}"
    );
    assert!(
        source.contains("rgb(192, 80, 77)"),
        "series 2 must take the theme's accent2, got:\n{source}"
    );
    assert!(
        !source.contains("rgb(68, 114, 196)"),
        "the built-in 2013+ accent1 must not appear when the file names its own, got:\n{source}"
    );
}

#[test]
fn test_automatic_series_colors_keep_the_builtin_palette_without_a_theme() {
    // Triangulation: a file that supplies no accents still renders, on the
    // built-in palette rather than on nothing.
    let source = chart_source(two_series_bar_chart(Vec::new()));

    assert!(
        source.contains("rgb(68, 114, 196)"),
        "the built-in palette stands in when the package names no accents, got:\n{source}"
    );
}

#[test]
fn test_explicit_series_fill_still_outranks_the_theme() {
    // Triangulation, and the guarantee #535 established: a fill the file
    // states wins over any automatic colour.
    let mut chart = two_series_bar_chart(office_2007_accents());
    chart.series[0].fill = Some(crate::ir::Color::new(0x11, 0x22, 0x33));
    let source = chart_source(chart);

    assert!(
        source.contains("rgb(17, 34, 51)"),
        "the declared fill must survive, got:\n{source}"
    );
    assert!(
        source.contains("rgb(192, 80, 77)"),
        "the series that declares none still takes accent2, got:\n{source}"
    );
}

#[test]
fn test_line_series_markers_cycle_by_series_index() {
    // `c:marker val="1"` with no `c:symbol` means "the default marker for this
    // series index", and the point of the sequence is that adjacent series stay
    // apart in monochrome. Drawing one square for every series defeats that
    // (issue #635).
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Line;
    chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
    chart.series[0].values = vec![100.0, 120.0];
    chart.series[1].values = vec![60.0, 80.0];
    let source = chart_source(chart);

    assert!(
        source.contains("polygon("),
        "a cycled marker set needs shapes beyond `rect`, got:\n{source}"
    );
    // Series 1 and series 2 must not draw the same marker.
    let squares = source.matches("rect(width: 5pt, height: 5pt").count();
    let polygons = source.matches("polygon(").count();
    assert!(
        squares > 0 && polygons > 0,
        "the two series must draw different marker shapes, got {squares} squares \
         and {polygons} polygons in:\n{source}"
    );
}

/// A line chart of two three-point series, each drawing `symbols`.
fn line_chart_with_markers(symbols: [Option<MarkerSymbol>; 2]) -> Chart {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Line;
    chart.categories = vec!["Q1".to_string(), "Q2".to_string(), "Q3".to_string()];
    chart.series[0].values = vec![4.0, 8.0, 6.0];
    chart.series[1].values = vec![6.0, 2.0, 5.0];
    for (series, symbol) in chart.series.iter_mut().zip(symbols) {
        series.marker_symbol = symbol;
    }
    chart
}

/// How many markers of each shape the source draws, as
/// `(circles, polygons, squares)`. A line chart draws nothing else as any of
/// them, so every hit is a marker — in the plot or on a legend key.
fn marker_shape_counts(source: &str) -> (usize, usize, usize) {
    let circle: String = format!(
        "circle(radius: {}pt",
        format_f64(SERIES_MARKER_SIZE_PT / 2.0)
    );
    let square: String = format!("rect(width: {}pt", format_f64(SERIES_MARKER_SIZE_PT));
    (
        source.matches(circle.as_str()).count(),
        source.matches("polygon(").count(),
        source.matches(square.as_str()).count(),
    )
}

#[test]
fn a_declared_marker_symbol_outranks_the_shape_cycle() {
    // A series that names its own `<c:marker><c:symbol>` gets that symbol
    // whatever its index. The cycle only ever stood in for a symbol the file
    // left automatic, so leaving it in charge drew the fourth series of the
    // audited workbook as a cross where Excel draws a filled circle (#1107).
    //
    // The two declarations here are each the shape the cycle would *not* have
    // picked: index 0 cycles to a diamond and index 1 to a square.
    let source = chart_source(line_chart_with_markers([
        Some(MarkerSymbol::Circle),
        Some(MarkerSymbol::Triangle),
    ]));

    // Three points per series, plus one marker on each series' legend key.
    let markers_per_series: usize = 3 + 1;
    let (circles, polygons, squares) = marker_shape_counts(&source);
    assert_eq!(
        circles, markers_per_series,
        "the series declaring `circle` must draw one on each point and on its \
         legend key; got:\n{source}"
    );
    assert_eq!(
        polygons, markers_per_series,
        "the series declaring `triangle` must draw one on each point and on its \
         legend key; got:\n{source}"
    );
    assert_eq!(
        squares, 0,
        "no series may fall back to the cycle's square; got:\n{source}"
    );
}

#[test]
fn a_marker_symbol_of_none_draws_no_marker_at_all() {
    // `<c:symbol val="none"/>` is a line series asking for a bare line. The
    // stroke and its legend key stay, so the series is still readable.
    let source = chart_source(line_chart_with_markers([
        Some(MarkerSymbol::Off),
        Some(MarkerSymbol::Off),
    ]));

    assert_eq!(
        marker_shape_counts(&source),
        (0, 0, 0),
        "a series whose symbol is `none` draws no point marker; got:\n{source}"
    );
    assert!(
        source.contains(&format!(
            "line(end: ({}pt, 0pt), stroke: {}pt",
            format_f64(LEGEND_KEY_LEN_PT),
            format_f64(SERIES_LINE_PT)
        )),
        "the legend key still samples the series line; got:\n{source}"
    );
}

/// The radius of the marker drawn on each line legend key, in emission
/// order. A key is the one `#box(width: 19.2pt, …)[…]` per entry, and the
/// marker inside it is the only `circle(` on that line.
fn legend_key_circle_radii(source: &str) -> Vec<f64> {
    let key_prefix: String = format!("#box(width: {}pt", format_f64(LEGEND_KEY_LEN_PT));
    source
        .lines()
        .filter(|line| line.contains(key_prefix.as_str()))
        .filter_map(|line| {
            let open: usize = line.find("circle(radius: ")?;
            let value: &str = &line[open + "circle(radius: ".len()..];
            let end: usize = value.find("pt")?;
            value[..end].parse::<f64>().ok()
        })
        .collect()
}

/// A one-series line chart of three points drawing `declared_marker_pt`
/// circles under a legend set at `legend_pt`.
fn line_chart_with_legend_size(legend_pt: f64, declared_marker_pt: f64) -> Chart {
    let mut chart = line_chart_with_markers([Some(MarkerSymbol::Circle), Some(MarkerSymbol::Off)]);
    chart.series[0].marker_style.size_pt = Some(declared_marker_pt);
    chart.legend_text_style.size_pt = Some(legend_pt);
    chart
}

#[test]
fn a_legend_marker_never_exceeds_six_tenths_of_the_legend_font_size() {
    // Excel for Mac 16.112, one factor per export on the #982 gift-budget
    // workbook with the line series declaring a 24pt marker: the legend key's
    // marker measures 3, 4, 4, 5, 6, 6, 7, 8 and 10 chart points at legend
    // sizes 6, 7, 8, 9, 10, 11, 12, 14 and 18 — floor(0.6 x size), with the
    // same values on a Verdana legend, whose painted key rectangle is a
    // different height. The plotted markers keep their declared 24pt (#1617).
    for (legend_pt, expected) in [
        (6.0, 3.0),
        (7.0, 4.0),
        (8.0, 4.0),
        (9.0, 5.0),
        (10.0, 6.0),
        (11.0, 6.0),
        (12.0, 7.0),
        (14.0, 8.0),
        (18.0, 10.0),
    ] {
        let source = chart_source(line_chart_with_legend_size(legend_pt, 24.0));

        assert_eq!(
            legend_key_circle_radii(&source),
            vec![expected / 2.0],
            "a {legend_pt}pt legend must draw its 24pt series marker at {expected}pt; got:\n{source}"
        );
        let plotted: usize = source.matches("circle(radius: 12pt").count();
        assert_eq!(
            plotted, 3,
            "the plot must keep one 24pt marker on each of its three points at a \
             {legend_pt}pt legend; got:\n{source}"
        );
    }
}

#[test]
fn a_declared_marker_below_the_legend_cap_keeps_its_size() {
    // The cap only limits. Native 9pt legend exports draw declared 2, 3 and 5pt
    // markers at 2, 3 and 5pt and clamp 7 and 12 to 5; a 6pt legend draws a
    // declared 2pt marker at 2pt and clamps 4, 5 and 7 to 3 (#1617).
    for (legend_pt, declared, expected) in [
        (9.0, 2.0, 2.0),
        (9.0, 3.0, 3.0),
        (9.0, 5.0, 5.0),
        (9.0, 7.0, 5.0),
        (9.0, 12.0, 5.0),
        (6.0, 2.0, 2.0),
        (6.0, 4.0, 3.0),
        (6.0, 5.0, 3.0),
        (6.0, 7.0, 3.0),
    ] {
        let source = chart_source(line_chart_with_legend_size(legend_pt, declared));

        assert_eq!(
            legend_key_circle_radii(&source),
            vec![expected / 2.0],
            "a {legend_pt}pt legend must draw a declared {declared}pt marker at \
             {expected}pt; got:\n{source}"
        );
        let plotted: usize = source
            .matches(&format!("circle(radius: {}pt", format_f64(declared / 2.0)))
            .count();
        assert!(
            plotted >= 3,
            "the plot must keep its declared {declared}pt marker on every point; got:\n{source}"
        );
    }
}

#[test]
fn a_small_legend_shrinks_the_line_sample_box_with_its_marker() {
    // At 6pt the native legend keeps all four labels on one baseline. A 5pt
    // sample box outgrows a 6pt line box and pushed the line series' label
    // 1.06 printed points below the column labels (#1617 work record), so the
    // box is no taller than the capped marker — and no shorter than today's
    // 5pt allowance at ordinary sizes, where the calibrated key seat stays.
    let source = chart_source(line_chart_with_legend_size(6.0, 5.0));
    assert!(
        source.contains(&format!(
            "#box(width: {}pt, height: 3pt, baseline: {}pt)[",
            format_f64(LEGEND_KEY_LEN_PT),
            format_f64(LEGEND_KEY_BASELINE_PT)
        )),
        "a 6pt legend's line sample box must be 3pt tall; got:\n{source}"
    );

    for declared in [2.0, 5.0, 12.0] {
        let source = chart_source(line_chart_with_legend_size(9.0, declared));
        assert!(
            source.contains(&format!(
                "#box(width: {}pt, height: {}pt, baseline: {}pt)[",
                format_f64(LEGEND_KEY_LEN_PT),
                format_f64(SERIES_MARKER_SIZE_PT),
                format_f64(LEGEND_KEY_BASELINE_PT)
            )),
            "a 9pt legend keeps the {}pt sample box for a declared {declared}pt marker; got:\n{source}",
            format_f64(SERIES_MARKER_SIZE_PT)
        );
    }
}

#[test]
fn the_legend_marker_cap_reaches_every_family_that_draws_a_line_key() {
    // The line key is emitted from three places: the line family, a line series
    // inside a bar/column plot area (#1067) and the radar family. A cap in any
    // one of them still leaves the others oversized.
    for (chart_type, plot_type, label) in [
        (ChartType::Line, None, "line"),
        (
            ChartType::Other(crate::ir::RADAR_CHART_LABEL.to_string()),
            None,
            "radar",
        ),
        (
            ChartType::Column,
            Some(ChartType::Line),
            "line over columns",
        ),
    ] {
        let mut chart = line_chart_with_legend_size(6.0, 5.0);
        chart.chart_type = chart_type;
        for series in chart.series.iter_mut() {
            series.plot_type = plot_type.clone();
        }
        let source = chart_source(chart);

        assert_eq!(
            legend_key_circle_radii(&source),
            vec![1.5],
            "{label}: a 6pt legend must draw the 5pt series marker at 3pt; got:\n{source}"
        );
    }
}

/// Every size a `#text(size: Npt)[label]` was emitted at, for one label.
fn emitted_text_sizes(source: &str, label: &str) -> Vec<f64> {
    let suffix: String = format!(")[{label}]");
    let mut sizes: Vec<f64> = Vec::new();
    for (index, _) in source.match_indices(&suffix) {
        let Some(open) = source[..index].rfind("#text(size: ") else {
            continue;
        };
        let value: &str = &source[open + "#text(size: ".len()..index];
        if let Ok(size) = value.trim_end_matches("pt").parse::<f64>() {
            sizes.push(size);
        }
    }
    sizes
}

#[test]
fn chart_labels_take_the_default_chart_text_size() {
    // A chart that declares no `c:txPr` anywhere still has a text size: Excel's
    // 10pt chart default. The sizes were per-element constants instead — 8pt for
    // the axis and category labels, 9pt for the legend — so the labels rendered
    // at a size the file never asks for, and did not even agree with each other
    // (issue #800).
    //
    // Both renderers that can be checked against `WithChart.xlsx` put every run
    // at 10pt: the native Excel export measures a 6.24pt cap height (10pt
    // Calibri) and LibreOffice writes a literal 10.0pt text matrix for all 18
    // runs on the page.
    for chart_type in [ChartType::Bar, ChartType::Line] {
        let mut chart = two_series_bar_chart(Vec::new());
        let kind: String = format!("{chart_type:?}");
        chart.chart_type = chart_type.clone();
        chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
        chart.series[0].values = vec![4.0, 8.0];
        chart.series[1].values = vec![6.0, 2.0];
        let source = chart_source(chart);

        for label in ["0", "Q1", "Q2", "Revenue", "Cost"] {
            let sizes = emitted_text_sizes(&source, label);
            assert!(
                !sizes.is_empty(),
                "{kind}: no #text(size:) wrapped the label {label}; got:\n{source}"
            );
            for size in &sizes {
                assert_eq!(
                    *size, CHART_DEFAULT_TEXT_PT,
                    "{kind}: label {label} drew at {size}pt, not the \
                     {CHART_DEFAULT_TEXT_PT}pt chart default; got:\n{source}"
                );
            }
        }
    }
}

#[test]
fn legend_keys_use_an_explicit_chart_owned_label_gap() {
    // A plain markup space inherits the document's body font and size. That
    // leaked an 11pt word-space run between each 10pt chart key and label,
    // widening the legend independently of the chart's own text (#804).
    for chart_type in [ChartType::Bar, ChartType::Line, ChartType::Pie] {
        let mut chart = two_series_bar_chart(Vec::new());
        let kind = format!("{chart_type:?}");
        chart.chart_type = chart_type.clone();
        chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
        chart.series[0].values = vec![4.0, 8.0];
        chart.series[1].values = vec![6.0, 2.0];
        let source = chart_source(chart);

        assert!(
            source.contains("#h(0pt)#text(size: 10pt)"),
            "{kind}: the key-to-label gap must be explicit chart layout; got:\n{source}"
        );
    }
}

#[test]
fn a_line_legend_key_draws_the_series_line_and_its_marker() {
    // Excel's legend key for a line series is a sample of what the reader sees
    // in the plot: the series line at its own weight with the series' marker
    // centred on it. A filled 12x3pt bar carries neither, so the key could not
    // be matched to its line (issue #801).
    //
    // Measured on the native export of `WithChart.xlsx`: the key line is
    // 20.16pt and 20.64pt long for the two series, and each carries a ~5pt
    // marker centred on it — a diamond for the first series, a square for the
    // second.
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Line;
    chart.categories = vec!["Q1".to_string(), "Q2".to_string(), "Q3".to_string()];
    chart.series[0].values = vec![4.0, 8.0, 6.0];
    chart.series[1].values = vec![6.0, 2.0, 5.0];
    let source = chart_source(chart);

    assert!(
        !source.contains("height: 3pt, fill:"),
        "the legend key must not be a flat filled bar; got:\n{source}"
    );
    // One marker per data point, plus one on each series' legend key.
    let points_per_series: usize = 3;
    for (shape, label) in [("polygon(", "diamond"), ("rect(width: 5pt", "square")] {
        assert_eq!(
            source.matches(shape).count(),
            points_per_series + 1,
            "the {label} series must draw a marker on its legend key as well as \
             on each of its {points_per_series} points; got:\n{source}"
        );
    }
    assert!(
        source.contains(&format!(
            "line(end: ({}pt, 0pt), stroke: {}pt",
            format_f64(LEGEND_KEY_LEN_PT),
            format_f64(SERIES_LINE_PT)
        )),
        "the legend key must draw the series line at its own weight; got:\n{source}"
    );
}

/// The weight of every polyline `curve(stroke: Npt + ...)` and radar
/// `polygon(stroke: Npt + ...)` the source strokes, in emission order.
fn emitted_path_stroke_widths(source: &str) -> Vec<f64> {
    let mut widths: Vec<f64> = Vec::new();
    // One `#place(…)` per line, so the line bounds each match and a shape
    // without a stroke cannot borrow the next one's.
    for line in source
        .lines()
        .filter(|line| line.contains("curve(stroke:") || line.contains("polygon(stroke:"))
    {
        let Some(open) = line.find("stroke: ") else {
            continue;
        };
        let value: &str = &line[open + "stroke: ".len()..];
        let Some(end) = value.find("pt ") else {
            continue;
        };
        if let Ok(width) = value[..end].parse::<f64>() {
            widths.push(width);
        }
    }
    widths
}

/// A line chart whose two series declare the weights in `widths`.
fn line_chart_with_line_widths(widths: [Option<f64>; 2]) -> Chart {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Line;
    chart.categories = vec!["Q1".to_string(), "Q2".to_string(), "Q3".to_string()];
    chart.series[0].values = vec![4.0, 8.0, 6.0];
    chart.series[1].values = vec![6.0, 2.0, 5.0];
    for (series, width) in chart.series.iter_mut().zip(widths) {
        series.line_width_pt = width;
    }
    chart
}

#[test]
fn a_series_polyline_takes_the_weight_its_line_declares() {
    // `<a:ln w="28440"/>` is 2.24pt, 12% heavier than the renderer's flat
    // 2.0pt. Excel prints the declared weight, so a workbook whose gridlines
    // agree to the point still had its plotted line print visibly thin
    // (issue #1113).
    let source = chart_source(line_chart_with_line_widths([Some(2.2394), None]));

    let widths: Vec<f64> = emitted_path_stroke_widths(&source);
    assert!(
        widths.contains(&2.2394),
        "the series declaring 28440 EMU must stroke at 2.2394pt; got {widths:?} in:\n{source}"
    );
    assert!(
        widths.contains(&SERIES_LINE_PT),
        "the series declaring no weight must keep the default {}pt; got {widths:?} in:\n{source}",
        format_f64(SERIES_LINE_PT)
    );
}

#[test]
fn a_legend_key_samples_its_series_declared_weight() {
    // The key stands for the line only while it is drawn at the line's own
    // weight (#801), so a declared weight has to reach both.
    let source = chart_source(line_chart_with_line_widths([Some(2.2394), None]));

    for (weight, which) in [
        (2.2394, "declaring one"),
        (SERIES_LINE_PT, "declaring none"),
    ] {
        assert!(
            source.contains(&format!(
                "line(end: ({}pt, 0pt), stroke: {}pt",
                format_f64(LEGEND_KEY_LEN_PT),
                format_f64(weight)
            )),
            "the legend key of the series {which} must sample it at {}pt; got:\n{source}",
            format_f64(weight)
        );
    }
}

#[test]
fn a_declared_weight_reaches_every_family_that_plots_a_line() {
    // The polyline is emitted from three places — the line family, a line
    // series inside a bar/column plot area (#1067), and the radar family's
    // closed polygon. A constant in any one of them still ignores the file.
    for (chart_type, plot_type, label) in [
        (ChartType::Line, None, "line"),
        (
            ChartType::Other(crate::ir::RADAR_CHART_LABEL.to_string()),
            None,
            "radar",
        ),
        (
            ChartType::Column,
            Some(ChartType::Line),
            "line over columns",
        ),
    ] {
        let mut chart = line_chart_with_line_widths([Some(2.2394), Some(1.5)]);
        chart.chart_type = chart_type.clone();
        for series in chart.series.iter_mut() {
            series.plot_type = plot_type.clone();
        }
        let source = chart_source(chart);

        let widths: Vec<f64> = emitted_path_stroke_widths(&source);
        for weight in [2.2394, 1.5] {
            assert!(
                widths.contains(&weight),
                "a {label} series declaring {weight}pt must stroke at it; \
                 got {widths:?} in:\n{source}"
            );
        }
    }
}

#[test]
fn every_chart_family_draws_the_default_chart_area_outline() {
    // A `c:chartSpace` with no `c:spPr/a:ln` still takes Office's default chart-area
    // outline — a thin rectangle enclosing the plot, the axis labels and the legend — so a
    // chart drawn without one has no boundary against the sheet behind it (#637).
    //
    // Measured on the native Excel export of `WithChart.xlsx` at 150 DPI: the border is a
    // single pixel of RGB(133,133,133), indistinguishable from the same page's gridlines,
    // which is `CHART_AUTOMATIC_LINE` — 0.75pt of #868686.
    for chart_type in [ChartType::Bar, ChartType::Line, ChartType::Pie] {
        let kind: String = format!("{chart_type:?}");
        let mut chart = two_series_bar_chart(Vec::new());
        chart.chart_type = chart_type.clone();
        chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
        chart.series[0].values = vec![4.0, 8.0];
        chart.series[1].values = vec![6.0, 2.0];
        let source = chart_source(chart);

        let outline: String = format!("stroke: {CHART_AREA_OUTLINE})[");
        assert!(
            source.contains(&outline),
            "{kind}: the chart area must carry the default outline; got:\n{source}"
        );
        // Exactly one box takes it — the outermost. A stroke on a nested box would draw a
        // second rectangle inside the chart.
        assert_eq!(
            source.matches(&outline).count(),
            1,
            "{kind}: only the chart-area box may carry the outline; got:\n{source}"
        );
    }
}

#[test]
fn every_chart_family_paints_a_declared_chart_area_fill_once() {
    for chart_type in [ChartType::Bar, ChartType::Line, ChartType::Pie] {
        let kind: String = format!("{chart_type:?}");
        let mut chart = two_series_bar_chart(Vec::new());
        chart.chart_type = chart_type.clone();
        chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
        chart.series[0].values = vec![4.0, 8.0];
        chart.series[1].values = vec![6.0, 2.0];
        chart.chart_area_fill = ChartAreaFill::Solid(crate::ir::Color::new(0x12, 0x34, 0x56));
        let source = chart_source(chart);

        assert_eq!(
            source.matches("fill: rgb(18, 52, 86)").count(),
            1,
            "{kind}: only the outermost chart-area box may carry the fill; got:\n{source}"
        );
    }
}

#[test]
fn a_titled_chart_paints_the_full_area_but_not_its_inner_content_box() {
    let mut chart = bar_chart_at(Some(14.0), &["Q1", "Q2"]);
    chart.title = Some("Sales".to_string());
    chart.chart_area_fill = ChartAreaFill::Solid(crate::ir::Color::new(0x12, 0x34, 0x56));
    let source = framed_chart_source(&chart, 321.0, 240.0);
    let area_start = "#box(width: 321pt, height: 240pt, fill: rgb(18, 52, 86), stroke:";
    let area_position = source
        .find(area_start)
        .unwrap_or_else(|| panic!("the full chart area must carry the fill, got:\n{source}"));
    let title_position = source
        .find("Sales")
        .unwrap_or_else(|| panic!("the title must be emitted, got:\n{source}"));

    assert!(
        area_position < title_position,
        "the fill must wrap the title"
    );
    assert_eq!(source.matches("fill: rgb(18, 52, 86)").count(), 1);
    assert!(
        source[title_position..].contains("fill: none, stroke: none"),
        "the nested content box must not repaint the area; got:\n{source}"
    );
}

#[test]
fn absent_and_explicitly_transparent_chart_area_fills_stay_transparent() {
    for fill in [ChartAreaFill::Unspecified, ChartAreaFill::Transparent] {
        let mut chart = two_series_bar_chart(Vec::new());
        chart.chart_area_fill = fill;
        let source = chart_source(chart);
        assert!(
            source.contains("fill: none, stroke:"),
            "{fill:?} must leave the chart area transparent; got:\n{source}"
        );
    }
}

#[test]
fn a_chart_that_asks_for_no_outline_gets_none() {
    // `<a:ln><a:noFill/></a:ln>` on `c:chartSpace/c:spPr` is the file saying it
    // wants no chart-area border. Drawing the default anyway puts a grey box
    // around every chart part that deliberately has none — `123233_charts.xlsx`
    // and `oxp_CU018-Chart-Cached-Data-41.pptx` among them (#637).
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Line;
    chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
    chart.series[0].values = vec![4.0, 8.0];
    chart.series[1].values = vec![6.0, 2.0];
    chart.chart_area_outline = ChartAreaOutline::Suppressed;
    let source = chart_source(chart);

    assert!(
        source.contains("stroke: none)["),
        "a suppressed outline must draw nothing; got:\n{source}"
    );
    assert!(
        !source.contains(&format!("stroke: {CHART_AREA_OUTLINE})[")),
        "the default outline must not override an explicit noFill; got:\n{source}"
    );
}

#[test]
fn a_chart_outline_keeps_its_own_width_and_colour() {
    // Chart parts declare lines of their own that the automatic grey is not:
    // `xlsx/office2pdf_repository_workbook.xlsx` a 9360 EMU #d9d9d9 one, and
    // `pptx/chart-picture-bg.pptx` a 28575 EMU accent one (#637).
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Line;
    chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
    chart.series[0].values = vec![4.0, 8.0];
    chart.series[1].values = vec![6.0, 2.0];
    chart.chart_area_outline = ChartAreaOutline::Explicit {
        width_pt: Some(0.7370079),
        color: Some(crate::ir::Color::new(0xd9, 0xd9, 0xd9)),
    };
    let source = chart_source(chart);

    assert!(
        source.contains("rgb(217, 217, 217)"),
        "the declared colour must reach the outline; got:\n{source}"
    );
    assert!(
        !source.contains(&format!("stroke: {CHART_AREA_OUTLINE})[")),
        "a declared line must not be replaced by the automatic one; got:\n{source}"
    );
}

// ----- The chart's declared text face (issue #668) -----

#[test]
fn chart_text_is_set_in_the_face_the_chart_declares() {
    // Every chart string used to fall through to the engine's default serif,
    // a face that appears nowhere else in the document. No sub-renderer names
    // a font, so one scoped `set` has to cover them all.
    let mut chart = two_series_bar_chart(Vec::new());
    chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
    chart.series[0].values = vec![4.0, 8.0];
    chart.series[1].values = vec![2.0, 6.0];
    chart.title = Some("Sales".to_string());
    chart.text_font_family = Some("Calibri".to_string());

    let source: String = chart_source(chart);
    assert!(
        source.contains("#set text(font: "),
        "the chart must set its declared face, got:\n{source}"
    );
    assert!(
        source.contains("Calibri"),
        "the declared face must reach the emitted font list, got:\n{source}"
    );
}

#[test]
fn a_chart_face_keeps_its_class_when_none_of_its_substitutes_are_installed() {
    // `Calibri` maps to Carlito and Liberation Sans, and either face may be
    // absent from a host. The exhausted chain then sent every chart string to
    // the engine's default serif (issue #1213).
    let mut chart = two_series_bar_chart(Vec::new());
    chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
    chart.series[0].values = vec![4.0, 8.0];
    chart.series[1].values = vec![2.0, 6.0];
    chart.title = Some("Sales".to_string());
    chart.text_font_family = Some("Calibri".to_string());

    let source: String = chart_source(chart);
    let set_line: &str = source
        .lines()
        .find(|line| line.starts_with("#set text(font: "))
        .expect("the chart sets a face");

    assert!(
        set_line.contains("\"Helvetica\"") || set_line.contains("\"DejaVu Sans\""),
        "the chain must end on a generic sans, not on the engine's serif: {set_line}"
    );
    assert!(
        !set_line.contains("Serif"),
        "a sans face must never gain a serif candidate: {set_line}"
    );
}

#[test]
fn a_chart_naming_no_face_sets_none() {
    // A chart whose package has no theme keeps the renderer's existing
    // behaviour rather than naming a face nothing resolves.
    let mut chart = two_series_bar_chart(Vec::new());
    chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
    chart.series[0].values = vec![4.0, 8.0];
    chart.series[1].values = vec![2.0, 6.0];
    chart.text_font_family = None;

    assert!(!chart_source(chart).contains("#set text(font: "));
}

#[test]
fn a_chart_with_korean_labels_keeps_an_east_asian_fallback() {
    // The declared face is Latin; the categories are not. A chain built from
    // the family alone would leave the Hangul to the engine's own pick.
    let mut chart = two_series_bar_chart(Vec::new());
    chart.categories = vec!["매출".to_string(), "비용".to_string()];
    chart.series[0].values = vec![4.0, 8.0];
    chart.series[1].values = vec![2.0, 6.0];
    chart.text_font_family = Some("Calibri".to_string());

    let source: String = chart_source(chart);
    let set_line: &str = source
        .lines()
        .find(|line| line.starts_with("#set text(font: "))
        .expect("the chart sets a face");
    assert!(
        set_line.contains(','),
        "a Korean chart needs a fallback chain, not a bare family: {set_line}"
    );
}

// ----- Run properties declared in c:txPr (issue #669) -----

fn sized_bar_chart(size_pt: f64) -> Chart {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
    chart.series[0].values = vec![4.0, 8.0];
    chart.series[1].values = vec![2.0, 6.0];
    chart.title = Some("Sales".to_string());
    chart.text_style = crate::ir::ChartTextStyle {
        size_pt: Some(size_pt),
        bold: None,
        letter_spacing_hundredths: None,
        color: None,
        ellipsis_overflow: false,
    };
    chart
}

#[test]
fn chart_labels_take_the_size_the_chart_declares() {
    // `bar-chart.pptx` asks for 18pt and rendered at 10 — a little over half
    // the size the file requested.
    let source: String = chart_source(sized_bar_chart(18.0));
    for label in ["Q1", "Q2"] {
        assert_eq!(
            emitted_text_sizes(&source, label),
            vec![18.0],
            "category label {label} must take the declared size, got:\n{source}"
        );
    }
    assert!(
        !source.contains("#text(size: 10pt)"),
        "nothing may fall back to the chart default once a size is declared:\n{source}"
    );
}

#[test]
fn a_chart_title_takes_office_s_scaled_size() {
    // Office renders the 18pt `bar-chart.pptx` declares as a 22pt title.
    let source: String = chart_source(sized_bar_chart(18.0));
    // `emitted_text_sizes` cannot read a `#text` carrying a weight, and the
    // title always carries one.
    assert!(
        source.contains("#text(size: 21.6pt, weight: \"bold\")[Sales]"),
        "the title must scale by 1.2, got:\n{source}"
    );
}

#[test]
fn a_chart_declaring_no_size_keeps_the_eleven_point_title() {
    // The default title size is what `AREA_TITLE_H` was measured against, so a
    // chart that declares nothing must not move.
    let mut chart = sized_bar_chart(18.0);
    chart.text_style = crate::ir::ChartTextStyle::default();
    assert!(chart_source(chart).contains("#text(size: 11pt, weight: \"bold\")[Sales]"));
}

// ----- A `c:title`'s own `c:txPr` (issue #1215) -----

fn title_run_style(
    size_pt: Option<f64>,
    bold: Option<bool>,
    color: Option<crate::ir::Color>,
) -> crate::ir::ChartTextStyle {
    crate::ir::ChartTextStyle {
        size_pt,
        bold,
        letter_spacing_hundredths: None,
        color,
        ellipsis_overflow: false,
    }
}

/// A chart whose space asks for 18pt and whose title states its own properties.
fn own_title_style_chart(title_style: crate::ir::ChartTextStyle) -> Chart {
    let mut chart = sized_bar_chart(18.0);
    chart.title_text_style = title_style;
    chart
}

/// A size the title states itself is the printed size, not something to scale.
///
/// `tests/fixtures/xlsx/any_sheets.xlsx` states `sz="1400"` on the title's own
/// `c:txPr`, and the Excel for Mac 16.100 export prints it at
/// `trm="14 0 0 14"`. The 1.2 factor belongs to the chart space's size, which
/// Office scales *into* a title size; a title size is already that.
#[test]
fn a_title_stating_its_own_size_prints_it_unscaled() {
    let source: String = chart_source(own_title_style_chart(title_run_style(
        Some(14.0),
        None,
        None,
    )));

    assert!(
        source.contains("#text(size: 14pt, weight: \"bold\")[Sales]"),
        "the title must take its own 14pt, got:\n{source}"
    );
    assert!(
        !source.contains("21.6pt"),
        "and must not scale the chart space's 18pt on top of it:\n{source}"
    );
}

/// The same rule at another size, so nothing can pass by returning 14.
#[test]
fn another_title_size_prints_unscaled_too() {
    let source: String = chart_source(own_title_style_chart(title_run_style(
        Some(9.0),
        None,
        None,
    )));

    assert!(
        source.contains("#text(size: 9pt, weight: \"bold\")[Sales]"),
        "the title must take its own 9pt, got:\n{source}"
    );
}

/// `b="0"` on the title's own run properties prints a regular title.
#[test]
fn a_title_stating_regular_weight_is_not_drawn_bold() {
    let source: String = chart_source(own_title_style_chart(title_run_style(
        Some(14.0),
        Some(false),
        None,
    )));

    assert!(
        source.contains("#text(size: 14pt)[Sales]"),
        "a title declaring b=\"0\" must print regular, got:\n{source}"
    );
}

/// `b="1"` prints bold, and so does a title that states no weight at all —
/// the bold every chart has always been drawn with stays the fallback.
#[test]
fn a_title_stating_no_weight_keeps_the_bold_it_had() {
    for stated in [None, Some(true)] {
        let source: String = chart_source(own_title_style_chart(title_run_style(
            Some(14.0),
            stated,
            None,
        )));
        assert!(
            source.contains("#text(size: 14pt, weight: \"bold\")[Sales]"),
            "b={stated:?} must still print bold, got:\n{source}"
        );
    }
}

/// When the title states no weight, the chart space's weight still governs it
/// before the renderer falls back to the legacy bold title.
#[test]
fn a_title_without_its_own_weight_inherits_the_chart_spaces_regular_weight() {
    let mut chart = own_title_style_chart(title_run_style(Some(14.0), None, None));
    chart.text_style.bold = Some(false);
    let source: String = chart_source(chart);

    assert!(
        source.contains("#text(size: 14pt)[Sales]"),
        "the chart-space b=\"0\" must make an unstated title regular, got:\n{source}"
    );
    assert!(
        !source.contains("#text(size: 14pt, weight: \"bold\")[Sales]"),
        "the legacy fallback must not outrank chart-space formatting:\n{source}"
    );
}

/// The title takes the colour its own `a:solidFill` states.
#[test]
fn a_title_stating_a_colour_is_drawn_in_it() {
    let source: String = chart_source(own_title_style_chart(title_run_style(
        Some(14.0),
        Some(false),
        Some(crate::ir::Color::new(0x59, 0x59, 0x59)),
    )));

    assert!(
        source.contains("#text(size: 14pt, fill: rgb(89, 89, 89))[Sales]"),
        "the title must take its declared grey, got:\n{source}"
    );
}

/// The band a stated title size takes, against native Excel.
///
/// Twelve Excel for Mac 16.100 exports of the `Chart` chartsheet of
/// `tests/fixtures/xlsx/any_sheets.xlsx`, forced to Letter landscape, with the
/// title's own `sz` rewritten one value at a time and nothing else touched.
/// Measured with `mutool draw -F trace` as the topmost major gridline's device
/// y less the chart area's own top edge, and with the title band separated out
/// by a thirteenth export whose `<c:title>` element is removed altogether —
/// that one starts its plot 11.00pt below the same edge, which is the plot's
/// own inset with no title above it at all:
///
/// | title `sz` | plot top below the chart area | less the 11.00pt inset |
/// | ---: | ---: | ---: |
/// | 7 | 32.09 | 21.09 |
/// | 8 | 33.83 | 22.83 |
/// | 9 | 35.56 | 24.56 |
/// | 10 | 37.28 | 26.28 |
/// | 11 | 39.02 | 28.02 |
/// | 12 | 40.75 | 29.75 |
/// | 14 | 44.20 | 33.20 |
/// | 16 | 47.66 | 36.66 |
/// | 18 | 51.12 | 40.12 |
/// | 24 | 61.50 | 50.50 |
/// | 32 | 75.33 | 64.33 |
/// | 36 | 82.24 | 71.24 |
///
/// A least-squares line over all twelve is 8.994 + 1.72912 em, and no export
/// is further than 0.007pt from it.
#[test]
fn a_stated_title_size_takes_the_band_excel_gives_it() {
    for (size_pt, band) in [(9.0, 24.56), (14.0, 33.20), (36.0, 71.24)] {
        let chart = own_title_style_chart(title_run_style(Some(size_pt), Some(false), None));
        let measured: f64 = chart_area_title_h(&chart);
        assert!(
            (measured - band).abs() < 0.01,
            "a {size_pt}pt title takes {band}pt, got {measured}pt"
        );
    }
}

/// A chartsheet title's baseline is seated from the chart area's top edge,
/// independently of whichever face Typst resolves for the title.
///
/// Twelve Excel for Mac 16.100 exports of the `Chart` chartsheet in
/// `tests/fixtures/xlsx/any_sheets.xlsx`, with only the title's `sz` changed,
/// fit `8.251pt + 1.26390em`; no measured baseline is further than 0.70pt from
/// that line (issue #1314).
#[test]
fn an_excel_chartsheet_title_takes_the_native_baseline_seat() {
    for (size_pt, baseline_pt) in [(9.0, 19.6261), (14.0, 25.9456), (36.0, 53.7514)] {
        let mut chart = own_title_style_chart(title_run_style(Some(size_pt), Some(false), None));
        chart.host = crate::ir::ChartHost::SpreadsheetChartsheet;
        let source: String = framed_chart_source(&chart, 480.0, 320.0);
        let title: &str = source
            .lines()
            .find(|line| line.contains("[Sales]"))
            .expect("the chart title is emitted");

        assert!(
            title.contains(&format!(
                "top-edge: {}pt, bottom-edge: \"baseline\"",
                format_f64(baseline_pt)
            )),
            "a {size_pt}pt chartsheet title seats its baseline {baseline_pt}pt below the chart-area top, got: {title}"
        );
    }
}

#[test]
fn an_anchored_excel_chart_keeps_its_existing_title_seat() {
    let mut chart = own_title_style_chart(title_run_style(Some(14.0), Some(false), None));
    chart.host = crate::ir::ChartHost::Spreadsheet;
    let source: String = framed_chart_source(&chart, 480.0, 320.0);
    let title: &str = source
        .lines()
        .find(|line| line.contains("[Sales]"))
        .expect("the chart title is emitted");

    assert!(
        title.contains("#align(center + horizon)[#text(size: 14pt)[Sales]]"),
        "an anchored worksheet chart has no measured chartsheet baseline rule: {title}"
    );
}

/// A chart whose title states no size of its own is untouched by all of this.
#[test]
fn a_title_stating_no_size_keeps_the_band_it_had() {
    let mut chart = sized_bar_chart(18.0);
    let before: f64 = chart_area_title_h(&chart);
    chart.title_text_style = title_run_style(None, Some(false), None);

    assert_eq!(
        chart_area_title_h(&chart),
        before,
        "the band follows a stated size, not a stated weight"
    );
}

#[test]
fn category_labels_take_the_axis_weight() {
    // `a:defRPr b="1"` on `c:catAx` was dropped, so bold category labels
    // rendered regular while the data labels beside them kept their own bold.
    let mut chart = sized_bar_chart(11.0);
    chart.category_axis_text_style = crate::ir::ChartTextStyle {
        size_pt: None,
        bold: Some(true),
        letter_spacing_hundredths: None,
        color: None,
        ellipsis_overflow: false,
    };
    let source: String = chart_source(chart);
    assert!(
        source.contains("#text(size: 11pt, weight: \"bold\")[Q1]"),
        "the category label must carry the axis' weight, got:\n{source}"
    );
}

#[test]
fn an_axis_size_overrides_the_chart_space_size_for_that_axis_only() {
    let mut chart = sized_bar_chart(18.0);
    chart.category_axis_text_style = crate::ir::ChartTextStyle {
        size_pt: Some(9.0),
        bold: None,
        letter_spacing_hundredths: None,
        color: None,
        ellipsis_overflow: false,
    };
    let source: String = chart_source(chart);
    assert_eq!(emitted_text_sizes(&source, "Q1"), vec![9.0]);
    // The title still follows the chart space.
    assert!(source.contains("#text(size: 21.6pt, weight: \"bold\")[Sales]"));
}

// ----- Radar charts (issue #679) -----

fn radar_chart() -> Chart {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Other(crate::ir::RADAR_CHART_LABEL.to_string());
    chart.categories = vec![
        "Deploy".to_string(),
        "Startup".to_string(),
        "Deps".to_string(),
        "Portable".to_string(),
        "Coverage".to_string(),
    ];
    chart.series[0].values = vec![5.0, 5.0, 5.0, 5.0, 3.0];
    chart.series[1].values = vec![2.0, 2.0, 1.0, 3.0, 5.0];
    chart.title = Some("Qualitative".to_string());
    chart
}

#[test]
fn a_radar_chart_draws_a_plot_rather_than_a_data_table() {
    // #544 replaced the silently dropped chart with a bordered rectangle
    // holding an italic caption and a table of the series values, so a slide
    // whose primary content was a radar still lost it.
    let source: String = chart_source(radar_chart());
    assert!(
        !source.contains("Radar Chart"),
        "the type-label caption belongs to the table fallback, got:\n{source}"
    );
    assert!(
        source.contains("polygon(stroke:"),
        "a radar is drawn as closed rings and polygons, got:\n{source}"
    );
}

#[test]
fn a_radar_draws_one_closed_polygon_per_series() {
    // Two series over five categories: five web rings plus two series rings.
    let source: String = chart_source(radar_chart());
    let closed: usize = source.matches("polygon(stroke:").count();
    assert!(
        closed > 2,
        "expected a ring per major unit plus one polygon per series, got {closed} in:\n{source}"
    );
    // Each series polygon carries the series stroke width; the web does not.
    let series_rings: usize = source
        .matches(&format!(
            "polygon(stroke: {}pt + ",
            format_f64(SERIES_LINE_PT)
        ))
        .count();
    assert_eq!(
        series_rings, 2,
        "one closed polygon per series, got:\n{source}"
    );
}

#[test]
fn a_radar_labels_every_category_and_keeps_its_title() {
    let source: String = chart_source(radar_chart());
    for category in ["Deploy", "Startup", "Deps", "Portable", "Coverage"] {
        assert!(
            source.contains(&format!("[{category}]")),
            "category {category} must be labelled, got:\n{source}"
        );
    }
    assert!(source.contains("[Qualitative]"), "got:\n{source}");
}

#[test]
fn a_radar_with_too_few_categories_keeps_the_table_fallback() {
    // Two spokes cannot close a ring, so the table still says more than a
    // degenerate plot would.
    let mut chart = radar_chart();
    chart.categories = vec!["Deploy".to_string(), "Startup".to_string()];
    chart.series[0].values = vec![5.0, 5.0];
    chart.series[1].values = vec![2.0, 2.0];
    assert!(chart_source(chart).contains("Radar Chart"));
}

#[test]
fn a_radar_with_no_positive_value_keeps_the_table_fallback() {
    let mut chart = radar_chart();
    chart.series[0].values = vec![0.0; 5];
    chart.series[1].values = vec![0.0; 5];
    assert!(chart_source(chart).contains("Radar Chart"));
}

// ----- Plot chrome sized from the text it holds (issue #706) -----

fn bar_chart_at(size_pt: Option<f64>, categories: &[&str]) -> Chart {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Bar;
    chart.categories = categories.iter().map(|c| (*c).to_string()).collect();
    chart.series[0].values = vec![4.0; categories.len()];
    chart.series[1].values = vec![2.0; categories.len()];
    chart.text_style = crate::ir::ChartTextStyle {
        size_pt,
        bold: None,
        letter_spacing_hundredths: None,
        color: None,
        ellipsis_overflow: false,
    };
    chart
}

#[test]
fn a_chart_declaring_no_size_keeps_its_chrome_where_it_was() {
    // The band constants were calibrated at the 10pt chart default, so scaling
    // from them has to be the identity there or every untouched chart moves.
    let chart = bar_chart_at(None, &["Q1", "Q2"]);
    assert_eq!(chart_tick_band_pt(&chart), TICK_GAP);
    assert_eq!(chart_category_band_pt(&chart), ROW);
    assert_eq!(chart_category_gutter_pt(&chart), LABEL_W + GAP);
}

#[test]
fn a_larger_declared_size_reserves_a_taller_tick_band() {
    // Native PowerPoint reserves 39.9817pt below the plot for an 18pt chart;
    // the band includes both a fixed base and a text-scaled component, so a
    // simple 1.8x scaling is still short.
    let chart = bar_chart_at(Some(18.0), &["Q1", "Q2"]);
    assert!(
        (chart_tick_band_pt(&chart) - 39.9817).abs() < 0.02,
        "an 18pt chart reserves PowerPoint's measured band, got {}",
        chart_tick_band_pt(&chart)
    );
    assert!(chart_tick_band_pt(&chart) > TICK_GAP);
}

#[test]
fn a_framed_bar_chart_reserves_powerpoint_measured_chrome_at_multiple_sizes() {
    // Native PowerPoint 16.112 exports of `bar-chart.pptx` with only
    // `c:chartSpace/c:txPr/a:defRPr@sz` changed. Each value is the plot's
    // left/top/right/bottom edge relative to the 480 x 320pt graphic frame.
    // Two sizes keep the regression test from fitting the original 18pt GT
    // with constants that fail as soon as the chart text changes.
    let measurements = [
        (12.0, (55.3186, 37.4150, 413.0499, 291.0732)),
        (18.0, (79.7209, 46.2050, 391.0825, 279.9383)),
    ];

    for (size_pt, expected) in measurements {
        let mut chart = bar_chart_at(Some(size_pt), &["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]);
        chart.series.truncate(1);
        chart.series[0].name = Some("Sales".to_string());
        chart.has_legend = true;
        chart.legend_position = LegendPosition::Right;
        chart.text_font_family = Some("Calibri".to_string());
        let actual = axis_plot_rect(&chart, (480.0, 320.0), true);
        let errors = [
            ("left", actual.0, expected.0),
            ("top", actual.1, expected.1),
            ("right", actual.2, expected.2),
            ("bottom", actual.3, expected.3),
        ]
        .map(|(axis, actual, expected)| (axis, actual, expected, (actual - expected).abs()));
        assert!(
            errors.iter().all(|(_, _, _, error)| *error <= 0.1),
            "{size_pt}pt chart edges: {errors:?}"
        );
    }
}

#[test]
fn a_powerpoint_right_legend_places_its_scaled_entry_at_multiple_sizes() {
    // Native PowerPoint 16.112 exports of the same 480 x 320pt chart frame.
    // These are the key's left edge relative to the frame, its size, and the
    // visible key-to-label gap. Five sizes prevent a one-off translation fitted
    // only to the 18pt #841 GT.
    let measurements = [
        (10.0, 441.4465, 5.4923, 2.3710),
        (12.0, 435.6760, 6.5926, 2.9213),
        (18.0, 418.4018, 9.8887, 4.5694),
        (24.0, 401.1207, 13.1827, 6.2164),
        (36.0, 366.5520, 19.7753, 9.5126),
    ];

    for (size_pt, expected_x, expected_key_size, expected_gap) in measurements {
        let mut chart = bar_chart_at(Some(size_pt), &["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]);
        chart.series.truncate(1);
        chart.series[0].name = Some("Sales".to_string());
        chart.has_legend = true;
        chart.legend_position = LegendPosition::Right;
        chart.host = crate::ir::ChartHost::Presentation;
        chart.text_font_family = Some("Calibri".to_string());

        let source = framed_chart_source(&chart, 480.0, 320.0);
        let actual_x = legend_entry_x(&source, "Sales");
        assert!(
            (actual_x - expected_x).abs() <= 0.1,
            "{size_pt}pt PowerPoint legend key starts at {actual_x}pt, expected {expected_x}pt; got:\n{source}"
        );
        let entry = source
            .lines()
            .find(|line| line.contains("box[#box") && line.contains("[Sales]]"))
            .expect("the chart emits its Sales legend entry");
        let (key_size, key_height, gap) = *emitted_legend_key_boxes(entry)
            .first()
            .expect("the legend entry draws its filled key");
        assert!(
            (key_size - expected_key_size).abs() <= 0.002
                && (key_height - expected_key_size).abs() <= 0.002,
            "{size_pt}pt PowerPoint key is {key_size}pt by {key_height}pt, native {expected_key_size}pt; got:\n{source}"
        );
        assert!(
            (gap - expected_gap).abs() <= 0.001,
            "{size_pt}pt PowerPoint key leaves {gap}pt before its label, native {expected_gap}pt; got:\n{source}"
        );
    }
}

/// Native PowerPoint 16.112 exports of `tests/fixtures/pptx/bar-chart.pptx`
/// with only the chart space's `<a:latin>` typeface patched, driven by
/// `scripts/probes/issue-1439-legend-key-face.json` and gated by its
/// layout-identical re-zip control. Each row is the source face, its
/// `hhea` ascent-plus-descent sum in the face's own em, the key side that
/// export draws at 18pt chart text, and the gap it leaves before the label.
/// Read off `mutool draw -F trace`; the control is the Calibri row, which the
/// theme's minor font supplies (#1439).
const POWERPOINT_LEGEND_KEY_FACE_EXPORTS: [(&str, f64, f64, f64); 7] = [
    ("Calibri", (1950.0 + 550.0) / 2048.0, 9.8887, 4.5694),
    ("Arial", (1854.0 + 434.0) / 2048.0, 9.0495, 4.1498),
    ("Times New Roman", (1825.0 + 443.0) / 2048.0, 8.9707, 4.1104),
    ("Courier New", (1705.0 + 615.0) / 2048.0, 9.1755, 4.2127),
    ("Georgia", (1878.0 + 449.0) / 2048.0, 9.2025, 4.2262),
    ("Trebuchet MS", (1923.0 + 455.0) / 2048.0, 9.4050, 4.3275),
    ("Verdana", (2059.0 + 430.0) / 2048.0, 9.8438, 4.5469),
];

/// The 18pt PowerPoint right legend of the probe series, in `family`.
fn powerpoint_legend_key_sweep_chart(family: &str, size_pt: f64) -> Chart {
    let mut chart = bar_chart_at(Some(size_pt), &["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]);
    chart.series.truncate(1);
    chart.series[0].name = Some("Sales".to_string());
    chart.has_legend = true;
    chart.legend_position = LegendPosition::Right;
    chart.host = crate::ir::ChartHost::Presentation;
    chart.text_font_family = Some(family.to_string());
    chart
}

#[test]
fn a_powerpoint_legend_key_sizes_itself_from_the_legend_face_line_box() {
    // Seven faces over one 18pt chart separate the face from the size: a key
    // fitted in Calibri ems alone draws every one of them 9.8874pt wide, where
    // the native exports span 8.9707..9.8887pt.
    let mut checked: Vec<&str> = Vec::new();
    for (family, source_line_box_em, native_side_pt, native_gap_pt) in
        POWERPOINT_LEGEND_KEY_FACE_EXPORTS
    {
        // A runner without the source face installed resolves a substitute
        // whose line box belongs to another design, and the native number
        // cannot be asserted against that. Calibri's metrics are calibrated
        // in-tree and reach every runner, so the sweep never empties out --
        // the assertion after the loop is what holds that.
        let Some((ascent_em, descent_em)) = chart_face_line_metrics_em(family, false) else {
            continue;
        };
        if (ascent_em + descent_em - source_line_box_em).abs() > 1e-6 {
            continue;
        }
        checked.push(family);

        let source = framed_chart_source(
            &powerpoint_legend_key_sweep_chart(family, 18.0),
            480.0,
            320.0,
        );
        let keys: Vec<(f64, f64, f64)> = emitted_legend_key_boxes(&source);
        assert_eq!(
            keys.len(),
            1,
            "the one series takes one key; got:\n{source}"
        );
        let (width_pt, height_pt, gap_pt) = keys[0];
        assert!(
            (width_pt - native_side_pt).abs() <= 0.002
                && (height_pt - native_side_pt).abs() <= 0.002,
            "{family}: PowerPoint draws a {native_side_pt}pt square key, got {width_pt}pt by {height_pt}pt in:\n{source}"
        );
        assert!(
            (gap_pt - native_gap_pt).abs() <= 0.002,
            "{family}: PowerPoint leaves {native_gap_pt}pt before the label, got {gap_pt}pt in:\n{source}"
        );
    }
    assert!(
        checked.contains(&"Calibri"),
        "the calibrated Calibri row must be checked on every runner; checked {checked:?}"
    );
}

#[test]
fn a_powerpoint_legend_key_scales_that_line_box_with_chart_text() {
    // The same probe holding the face at Arial and moving the chart space's
    // `<a:defRPr sz>` instead
    // (`scripts/probes/issue-1439-legend-key-arial-size.json`): one line box
    // times the size, with no fixed term anywhere in the series.
    let (arial, arial_line_box_em, _, _) = POWERPOINT_LEGEND_KEY_FACE_EXPORTS[1];
    // Arial is the runner-dependent half of the sweep above; where the source
    // face is absent there is nothing native to compare a substitute against.
    let Some((ascent_em, descent_em)) = chart_face_line_metrics_em(arial, false) else {
        return;
    };
    if (ascent_em + descent_em - arial_line_box_em).abs() > 1e-6 {
        return;
    }
    for (size_pt, native_side_pt) in [
        (10.0_f64, 5.0288_f64),
        (12.0, 6.0345),
        (18.0, 9.0495),
        (24.0, 12.0646),
        (36.0, 18.0989),
    ] {
        let source = framed_chart_source(
            &powerpoint_legend_key_sweep_chart(arial, size_pt),
            480.0,
            320.0,
        );
        let keys: Vec<(f64, f64, f64)> = emitted_legend_key_boxes(&source);
        assert_eq!(
            keys.len(),
            1,
            "the one series takes one key; got:\n{source}"
        );
        assert!(
            (keys[0].0 - native_side_pt).abs() <= 0.002,
            "{size_pt}pt {arial}: PowerPoint draws a {native_side_pt}pt key, got {}pt in:\n{source}",
            keys[0].0
        );
    }
}

#[test]
fn an_unmeasurable_powerpoint_right_legend_keeps_the_plot_relative_fallback() {
    let mut chart = bar_chart_at(Some(18.0), &["Q1", "Q2"]);
    chart.series.truncate(1);
    chart.series[0].name = Some("Sales".to_string());
    chart.host = crate::ir::ChartHost::Presentation;
    chart.text_font_family = Some("Definitely Missing Chart Face 999".to_string());

    let plot_right = axis_plot_rect(&chart, (480.0, 320.0), false).2;
    let source = framed_chart_source(&chart, 480.0, 320.0);
    let actual_x = legend_entry_x(&source, "Sales");
    assert!(
        (actual_x - (plot_right + GAP)).abs() <= 0.01,
        "an unmeasurable face must preserve the plot-relative fallback, got {actual_x} after a {plot_right}pt plot; source:\n{source}"
    );
}

#[test]
fn a_powerpoint_right_legend_uses_the_native_vertical_center_at_multiple_sizes() {
    // Native PowerPoint 16.112 exports of the same 480 x 320pt chart frame.
    // Each value is the legend key's top edge inside the post-title chart body.
    // The native absolute edge is translated by the source frame and the same
    // title band used by `generate_chart_in`.
    let measurements = [
        (10.0, 133.7634),
        (12.0, 131.5833),
        (18.0, 125.0449),
        (24.0, 118.5077),
        (36.0, 105.4309),
    ];

    for (size_pt, expected_y) in measurements {
        let mut chart = bar_chart_at(Some(size_pt), &["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]);
        chart.series.truncate(1);
        chart.series[0].name = Some("Sales".to_string());
        chart.host = crate::ir::ChartHost::Presentation;
        chart.text_font_family = Some("Calibri".to_string());

        let source = framed_chart_source(&chart, 480.0, 320.0);
        let actual_y = legend_entry_y(&source, "Sales");
        assert!(
            (actual_y - expected_y).abs() <= 0.1,
            "{size_pt}pt PowerPoint legend key starts at y={actual_y}pt, expected {expected_y}pt; got:\n{source}"
        );
    }
}

#[test]
fn a_powerpoint_column_right_legend_uses_the_native_text_scaled_row_pitch() {
    // Page 8 of the #1407 PowerPoint fixture declares 11.97pt legend text.
    // Native PowerPoint places the three baselines 19.2pt apart, while the
    // legacy axis-legend constant compresses every pair to 14pt (#1434).
    let mut chart = bar_chart_at(Some(11.97), &["Year 1", "Year 2", "Net Profit"]);
    chart.chart_type = ChartType::Column;
    chart.host = crate::ir::ChartHost::Presentation;
    chart.legend_position = LegendPosition::Right;
    chart.text_font_family = Some("Avenir Next LT Pro".to_string());
    chart.series[0].name = Some("Total Sales".to_string());
    chart.series[1].name = Some("Total Cogs".to_string());
    let mut third = chart.series[0].clone();
    third.name = Some("Net Profit".to_string());
    chart.series.push(third);

    let source = framed_chart_source(&chart, 852.0, 280.0);
    let origins = [
        legend_entry_y(&source, "Total Sales"),
        legend_entry_y(&source, "Total Cogs"),
        legend_entry_y(&source, "Net Profit"),
    ];
    let pitches = [origins[1] - origins[0], origins[2] - origins[1]];
    assert!(
        pitches.iter().all(|pitch| (*pitch - 19.2).abs() <= 0.1),
        "11.97pt PowerPoint column legend row pitches are {pitches:?}, expected native 19.2pt; got:\n{source}"
    );

    chart.chart_type = ChartType::Bar;
    let bar_source = framed_chart_source(&chart, 852.0, 280.0);
    let bar_origins = [
        legend_entry_y(&bar_source, "Total Sales"),
        legend_entry_y(&bar_source, "Total Cogs"),
        legend_entry_y(&bar_source, "Net Profit"),
    ];
    let bar_pitches = [
        bar_origins[1] - bar_origins[0],
        bar_origins[2] - bar_origins[1],
    ];
    assert!(
        bar_pitches
            .iter()
            .all(|pitch| (*pitch - 14.0).abs() <= 0.01),
        "the separately calibrated PowerPoint bar path must keep its 14pt pitch, got {bar_pitches:?}; source:\n{bar_source}"
    );
}

#[test]
fn a_powerpoint_column_right_legend_uses_the_native_vertical_center_at_multiple_sizes() {
    // Native PowerPoint 16.112 exports of `tests/fixtures/pptx/bar-chart.pptx`
    // repackaged as a column chart, one `<c:defRPr sz>` at a time, through
    // `scripts/probes/issue-1435-column-legend-center.json`. PowerPoint puts
    // the right legend at the same page position for both families — the
    // 18pt column variant's key lands on the unpatched bar control's
    // `y[282.7380, 292.6268]` to the last emitted digit — so these targets are
    // the bar values of
    // `a_powerpoint_right_legend_uses_the_native_vertical_center_at_multiple_sizes`
    // re-measured on the column exports. Our own placement differs because a
    // column's content box ends in the category band rather than the value
    // tick band, which is what this correction removes (#1435).
    //
    // Each value is the key's top edge inside the post-title chart body. #1437
    // moved the plot's own top inset inside that body, and
    // `powerpoint_right_legend_y_shift` gives that movement back so these
    // native anchors do not follow it.
    let measurements = [
        (10.0, 133.7634),
        (12.0, 131.5808),
        (18.0, 125.0423),
        (24.0, 118.5052),
    ];

    for (size_pt, expected_y) in measurements {
        let mut chart = bar_chart_at(Some(size_pt), &["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]);
        chart.chart_type = ChartType::Column;
        chart.series.truncate(1);
        chart.series[0].name = Some("Sales".to_string());
        chart.host = crate::ir::ChartHost::Presentation;
        chart.text_font_family = Some("Calibri".to_string());

        let source = framed_chart_source(&chart, 480.0, 320.0);
        let actual_y = legend_entry_y(&source, "Sales");
        assert!(
            (actual_y - expected_y).abs() <= 0.01,
            "{size_pt}pt PowerPoint column legend key starts at y={actual_y}pt, expected {expected_y}pt; got:\n{source}"
        );
    }
}

#[test]
fn a_powerpoint_column_right_legend_correction_translates_the_whole_stack() {
    // The correction moves the block, so it must not disturb what the block is
    // made of: the #1434 row pitch stays, and legends of different lengths stay
    // centred on one another. Page 8 of the #1220 deck is the shape this
    // reproduces — 11.97pt legend text on a 690 x 220.8pt frame — where the
    // native three keys centre on 395.734pt and the uncorrected stack sat
    // 6.95pt below it.
    let column_legend = |names: &[&str]| -> Vec<f64> {
        let mut chart = bar_chart_at(Some(11.97), &["Year 1", "Year 2", "Year 3"]);
        chart.chart_type = ChartType::Column;
        chart.host = crate::ir::ChartHost::Presentation;
        chart.legend_position = LegendPosition::Right;
        chart.text_font_family = Some("Avenir Next LT Pro".to_string());
        chart.series.truncate(1);
        let template = chart.series[0].clone();
        chart.series = names
            .iter()
            .map(|name| {
                let mut series = template.clone();
                series.name = Some((*name).to_string());
                series
            })
            .collect();
        let source = framed_chart_source(&chart, 690.0, 220.8);
        names
            .iter()
            .map(|name| legend_entry_y(&source, name))
            .collect()
    };

    let row_h: f64 = 1.6 * 11.97;
    let three = column_legend(&["Total Sales", "Total Cogs", "Net Profit"]);
    let pitches = [three[1] - three[0], three[2] - three[1]];
    assert!(
        pitches.iter().all(|pitch| (*pitch - row_h).abs() <= 0.01),
        "the #1434 row pitch must survive the block correction, got {pitches:?}"
    );

    // A shorter legend has to centre on the same line, or the correction is
    // changing the stack rather than translating it.
    let two = column_legend(&["Total Sales", "Total Cogs"]);
    let centre_of =
        |origins: &[f64]| -> f64 { origins[0] + (origins.len() - 1) as f64 * row_h / 2.0 };
    assert!(
        (centre_of(&three) - centre_of(&two)).abs() <= 0.01,
        "two and three entries must share a centre, got {} and {}",
        centre_of(&two),
        centre_of(&three)
    );
}

#[test]
fn only_a_powerpoint_column_right_legend_takes_the_column_block_correction() {
    // The correction is a PowerPoint automatic-layout figure measured on
    // column exports, so it must not reach an Excel column chart, a legend on
    // another edge, or the separately calibrated PowerPoint bar path.
    let shift_of = |host: crate::ir::ChartHost,
                    chart_type: ChartType,
                    position: LegendPosition,
                    size_pt: f64|
     -> f64 {
        let mut chart = bar_chart_at(Some(size_pt), &["1st Qtr", "2nd Qtr"]);
        chart.chart_type = chart_type;
        chart.host = host;
        chart.legend_position = position;
        powerpoint_right_legend_y_shift(&chart, Some((480.0, 320.0)), 0.0)
    };

    // Read with no title band, so the correction is the fitted pair alone;
    // `a_powerpoint_column_right_legend_uses_the_native_vertical_center_at_multiple_sizes`
    // asserts the native key tops it lands on once #1437's measured band has
    // moved the box under it.
    for (size_pt, expected) in [
        (10.0, -6.5282),
        (12.0, -6.2498),
        (18.0, -5.4147),
        (24.0, -4.5796),
    ] {
        let actual = shift_of(
            crate::ir::ChartHost::Presentation,
            ChartType::Column,
            LegendPosition::Right,
            size_pt,
        );
        assert!(
            (actual - expected).abs() <= 0.001,
            "{size_pt}pt PowerPoint column right legend rises {actual}pt, expected {expected}pt"
        );
    }

    for (host, chart_type, position) in [
        (
            crate::ir::ChartHost::Spreadsheet,
            ChartType::Column,
            LegendPosition::Right,
        ),
        (
            crate::ir::ChartHost::Presentation,
            ChartType::Column,
            LegendPosition::TopRight,
        ),
        (
            crate::ir::ChartHost::Presentation,
            ChartType::Column,
            LegendPosition::Bottom,
        ),
    ] {
        let actual = shift_of(host, chart_type.clone(), position, 18.0);
        assert_eq!(
            actual, 0.0,
            "{host:?} {chart_type:?} {position:?} must keep the generic side stack, got {actual}pt"
        );
    }

    let bar = shift_of(
        crate::ir::ChartHost::Presentation,
        ChartType::Bar,
        LegendPosition::Right,
        18.0,
    );
    assert!(
        (bar - (PPTX_RIGHT_LEGEND_Y_SHIFT_PT + PPTX_RIGHT_LEGEND_Y_SHIFT_EM * 18.0)).abs() < 1e-9,
        "the horizontal-bar calibration must be untouched, got {bar}pt"
    );
}

#[test]
fn a_powerpoint_column_right_legend_right_fits_the_chart_frame_like_a_bar_chart() {
    // Native PowerPoint 16.112 exports of `tests/fixtures/pptx/bar-chart.pptx`
    // repackaged as a column chart, one `<c:defRPr sz>` at a time, through
    // `scripts/probes/issue-1435-column-legend-center.json`. Each value is the
    // legend key's left edge inside the 480 x 320pt chart frame, read off
    // `mutool draw -F trace`. Every one of them lands on the unpatched bar
    // control's x to the last emitted digit, so PowerPoint right-fits the
    // stack against the chart area rather than against the plot: the plot's
    // own right inset must not move the legend (#1436).
    let measurements = [
        (10.0, 441.4465),
        (12.0, 435.6760),
        (18.0, 418.4018),
        (24.0, 401.1207),
        (36.0, 366.5520),
    ];

    for (size_pt, expected_x) in measurements {
        let mut chart = bar_chart_at(Some(size_pt), &["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]);
        chart.chart_type = ChartType::Column;
        chart.series.truncate(1);
        chart.series[0].name = Some("Sales".to_string());
        chart.has_legend = true;
        chart.legend_position = LegendPosition::Right;
        chart.host = crate::ir::ChartHost::Presentation;
        chart.text_font_family = Some("Calibri".to_string());

        let column_x = legend_entry_x(&framed_chart_source(&chart, 480.0, 320.0), "Sales");
        assert!(
            (column_x - expected_x).abs() <= 0.1,
            "{size_pt}pt PowerPoint column legend key starts at {column_x}pt, expected the native {expected_x}pt"
        );

        // The same export measured the two families onto one x, so the
        // calibrated bar path is the second, independent statement of the rule.
        chart.chart_type = ChartType::Bar;
        let bar_x = legend_entry_x(&framed_chart_source(&chart, 480.0, 320.0), "Sales");
        assert!(
            (column_x - bar_x).abs() <= 0.01,
            "{size_pt}pt column legend sits at {column_x}pt against the bar family's {bar_x}pt, but PowerPoint puts both on one x"
        );
    }
}

#[test]
fn a_powerpoint_right_legend_follows_the_chart_frames_right_edge() {
    // The anchor the native exports name is the chart area's right edge, so
    // widening the frame has to carry the legend the whole way with it — for
    // both families, and whatever the plot reserves inside that edge. This
    // guards the anchor itself, which no single-frame measurement pins (#1436).
    let key_x = |chart_type: ChartType, frame_w: f64| -> f64 {
        let mut chart = bar_chart_at(Some(11.97), &["Year 1", "Year 2", "Year 3"]);
        chart.chart_type = chart_type;
        chart.series.truncate(1);
        chart.series[0].name = Some("Total Sales".to_string());
        chart.has_legend = true;
        chart.legend_position = LegendPosition::Right;
        chart.host = crate::ir::ChartHost::Presentation;
        chart.text_font_family = Some("Avenir Next LT Pro".to_string());
        legend_entry_x(&framed_chart_source(&chart, frame_w, 220.8), "Total Sales")
    };

    for chart_type in [ChartType::Column, ChartType::Bar] {
        for (narrow, wide) in [(480.0_f64, 617.0_f64), (617.0, 690.0)] {
            let moved = key_x(chart_type.clone(), wide) - key_x(chart_type.clone(), narrow);
            assert!(
                (moved - (wide - narrow)).abs() <= 0.01,
                "{chart_type:?} legend moved {moved}pt when the frame grew {}pt",
                wide - narrow
            );
        }
    }
}

#[test]
fn a_powerpoint_horizontal_value_axis_keeps_native_label_gap_at_multiple_sizes() {
    // Native PowerPoint 16.112 exports of the same 480 x 320pt chart frame.
    // Each value is the required Typst box-top gap after translating the
    // native zero glyph's top edge through Typst's size-scaled ink overhang.
    // Five sizes prevent an 18pt-only translation.
    let measurements = [
        (10.0, 11.3134),
        (12.0, 12.7448),
        (18.0, 17.2527),
        (24.0, 21.2363),
        (36.0, 30.5978),
    ];

    for (size_pt, expected_gap) in measurements {
        let mut chart = bar_chart_at(Some(size_pt), &["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]);
        chart.series.truncate(1);
        chart.series[0].name = Some("Sales".to_string());
        chart.host = crate::ir::ChartHost::Presentation;

        let source = framed_chart_source(&chart, 480.0, 320.0);
        let plot_bottom =
            axis_plot_rect(&chart, (480.0, 320.0), true).3 - chart_area_title_h(&chart);
        let label_top = horizontal_value_axis_label_y(&source, "0");
        let actual_gap = label_top - plot_bottom;
        assert!(
            (actual_gap - expected_gap).abs() <= 0.4,
            "{size_pt}pt PowerPoint zero label starts {actual_gap}pt below the axis, expected {expected_gap}pt; got:\n{source}"
        );
    }
}

#[test]
fn an_excel_worksheet_horizontal_value_axis_scales_its_label_gap_with_size() {
    // Native Excel for Mac 16.112 exports of the `issue_1181_fit_to_height.xlsx`
    // income bar chart (#1622), isolated to just the value-axis font size via
    // the one-factor `office`-backend probes in `assets/validation/issue-1622/`.
    // 8pt and 10pt are the only clean points -- native re-lays the plot itself
    // past that, so no third point could confirm the fit is linear rather than
    // merely a two-point line. Expected gaps translate the native measurement
    // through the same Typst text box's own top-to-baseline offset for this
    // Trebuchet MS face (see the constants' doc comment for the derivation).
    let measurements = [(8.0, 8.131547), (10.0, 8.656938)];

    for (size_pt, expected_gap) in measurements {
        let mut chart = monthly_budget_income_chart();
        chart.value_axis_text_style.size_pt = Some(size_pt);
        let source = framed_chart_source(
            &chart,
            MONTHLY_BUDGET_CHART_FRAME.0,
            MONTHLY_BUDGET_CHART_FRAME.1,
        );
        let plot_bottom = axis_plot_rect(&chart, MONTHLY_BUDGET_CHART_FRAME, false).3;
        let actual_gap = horizontal_value_axis_label_y(&source, "0") - plot_bottom;
        assert!(
            (actual_gap - expected_gap).abs() <= 0.01,
            "{size_pt}pt worksheet value-label gap is {actual_gap}pt, expected {expected_gap}pt; got:\n{source}"
        );
    }
}

#[test]
fn non_worksheet_horizontal_value_axes_keep_their_existing_label_gap() {
    for host in [
        crate::ir::ChartHost::SpreadsheetChartsheet,
        crate::ir::ChartHost::WordProcessing,
    ] {
        let mut chart = bar_chart_at(Some(18.0), &["Q1", "Q2"]);
        chart.host = host;
        let source = framed_chart_source(&chart, 480.0, 320.0);
        let plot_bottom = axis_plot_rect(&chart, (480.0, 320.0), false).3;
        let actual_gap = horizontal_value_axis_label_y(&source, "0") - plot_bottom;
        assert!(
            (actual_gap - 4.0).abs() <= 0.01,
            "the {host:?} value-label gap changed to {actual_gap}pt; got:\n{source}"
        );
    }
}

#[test]
fn a_framed_column_chart_reserves_powerpoint_measured_chrome() {
    // Native PowerPoint 16.112 export of slide 14 in the #841 Contoso deck.
    // The coordinates are relative to its 401.95 x 344.25pt graphic frame.
    let mut chart = crowded_column_chart();
    chart.text_style.size_pt = Some(11.97);
    chart.category_axis_text_style.size_pt = Some(11.97);
    chart.value_axis_text_style.size_pt = Some(11.97);
    chart.text_font_family = Some("Avenir Next LT Pro".to_string());
    chart.has_legend = false;
    let actual = axis_plot_rect(&chart, (401.95, 344.25), false);
    let expected = (46.9766, 12.266, 390.9504, 193.1674);
    let errors = [
        ("left", actual.0, expected.0),
        ("top", actual.1, expected.1),
        ("right", actual.2, expected.2),
        ("bottom", actual.3, expected.3),
    ]
    .map(|(axis, actual, expected)| (axis, actual, expected, (actual - expected).abs()));
    assert!(
        errors.iter().all(|(_, _, _, error)| *error <= 0.1),
        "column chart edges: {errors:?}"
    );
}

/// `tests/fixtures/pptx/bar-chart.pptx` repackaged as the column chart the
/// #1437 probes export, on the deck's own 480 x 320pt graphic frame.
///
/// `chart_space_pt` is the `c:chartSpace/c:txPr` size every string inherits;
/// `title_pt` is the size the title states for itself, which PowerPoint prints
/// instead of scaling the chart space's by [`CHART_AREA_TITLE_SCALE`].
fn powerpoint_column_probe_chart(chart_space_pt: f64, title_pt: Option<f64>) -> Chart {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Column;
    chart.host = crate::ir::ChartHost::Presentation;
    chart.series.truncate(1);
    chart.series[0].name = Some("Sales".to_string());
    chart.series[0].values = vec![8.2, 3.2, 1.4, 1.2];
    chart.categories = ["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]
        .iter()
        .map(|category| (*category).to_string())
        .collect();
    chart.title = Some("Sales".to_string());
    chart.text_style.size_pt = Some(chart_space_pt);
    chart.title_text_style.size_pt = title_pt;
    chart
}

/// Plot top and bottom of the probe chart, in points below the frame's own
/// top edge and above its bottom edge.
fn powerpoint_column_probe_bands(chart: &Chart) -> (f64, f64) {
    let (_, top, _, bottom) = axis_plot_rect(chart, (480.0, 320.0), true);
    (top, 320.0 - bottom)
}

#[test]
fn a_powerpoint_column_plot_takes_the_native_automatic_top_band() {
    // Native PowerPoint 16.112 exports of `bar-chart.pptx` repackaged as a
    // column chart on its 480 x 320pt frame, plot top read off the value axis
    // line with `mutool draw -F trace`, in points below the frame's top edge.
    //
    // `scripts/probes/issue-1435-column-legend-center.json` moves the chart
    // space's own size, which carries the automatic title with it at 1.2x.
    for (chart_space_pt, native_top) in [
        (10.0, 34.752),
        (12.0, 38.900),
        (18.0, 51.353),
        (24.0, 63.802),
        (36.0, 88.708),
    ] {
        let chart = powerpoint_column_probe_chart(chart_space_pt, None);
        let (top, _) = powerpoint_column_probe_bands(&chart);
        assert!(
            (top - native_top).abs() <= 0.05,
            "a {chart_space_pt}pt chart space must start its plot {native_top}pt \
             below the frame; got {top}"
        );
    }
    // `scripts/probes/issue-1437-column-plot-top.json` holds the chart space at
    // 18pt and states the title's own size instead, so the two series separate
    // the title band from the plot's own inset.
    for (title_pt, native_top) in [
        (10.0, 37.192),
        (14.0, 42.077),
        (18.0, 46.962),
        (24.0, 54.282),
        (36.0, 68.932),
    ] {
        let chart = powerpoint_column_probe_chart(18.0, Some(title_pt));
        let (top, _) = powerpoint_column_probe_bands(&chart);
        assert!(
            (top - native_top).abs() <= 0.05,
            "a {title_pt}pt title must start the plot {native_top}pt below the \
             frame; got {top}"
        );
    }
}

#[test]
fn a_powerpoint_column_plot_reserves_the_native_category_band() {
    // The same exports, plot bottom read off the category axis line. The 36pt
    // chart space is left out: PowerPoint wraps `4th Qtr` onto a second line
    // there and we slant it instead, so neither side is drawing one flat band
    // (issue #1675).
    for (chart_space_pt, native_band) in [
        (10.0, 25.052),
        (12.0, 28.767),
        (18.0, 39.902),
        (24.0, 51.028),
    ] {
        let chart = powerpoint_column_probe_chart(chart_space_pt, None);
        let (_, band) = powerpoint_column_probe_bands(&chart);
        assert!(
            (band - native_band).abs() <= 0.1,
            "a {chart_space_pt}pt category axis must reserve {native_band}pt \
             under the plot; got {band}"
        );
    }
    // The band holds the category labels, so the title's size cannot move it.
    for title_pt in [10.0, 14.0, 18.0, 24.0, 36.0] {
        let chart = powerpoint_column_probe_chart(18.0, Some(title_pt));
        let (_, band) = powerpoint_column_probe_bands(&chart);
        assert!(
            (band - 39.902).abs() <= 0.1,
            "a {title_pt}pt title must leave the 18pt category band at 39.902pt; \
             got {band}"
        );
    }
}

#[test]
fn a_powerpoint_bar_plot_keeps_its_own_measured_top_band() {
    // The unpatched re-zip control of both probes is the same deck as a bar
    // chart: 18pt chart space, 22pt automatic title, plot top 46.365pt below
    // the frame. The column calibration must not reach the bar family, whose
    // band [`CHART_PLOT_TOP_PAD_PT`] still carries the top inset a bar plot
    // keeps on its own account.
    let mut chart = powerpoint_column_probe_chart(18.0, None);
    chart.chart_type = ChartType::Bar;
    let (top, _) = powerpoint_column_probe_bands(&chart);
    assert!(
        (top - 46.365).abs() <= 0.2,
        "the bar control must keep its 46.365pt top band; got {top}"
    );
}

#[test]
fn an_explicit_powerpoint_column_value_axis_keeps_the_native_label_inset() {
    // The #841 chart's plot is already aligned, but each right-aligned tick
    // label was 6.087pt left of PowerPoint. The calibrated plot gutter still
    // needs the same 6pt inner inset the legacy `TICK_GAP + GAP` layout had.
    let mut chart = crowded_column_chart();
    chart.host = crate::ir::ChartHost::Presentation;
    chart.text_style.size_pt = Some(11.97);
    chart.value_axis_text_style.size_pt = Some(11.97);
    let source = framed_chart_source(&chart, 401.95, 344.25);
    let zero = source
        .lines()
        .find(|line| line.contains("align(right + horizon)") && line.ends_with("[0]]])"))
        .expect("the zero value-axis label is emitted");
    assert!(zero.contains("dx: 6pt"), "{zero}");
    assert!(
        zero.contains("box(width: 28.784350000000003pt"),
        "the fix must translate, not widen, the value-label box: {zero}"
    );
}

/// The #1166 workbook's anchored column chart, restated in Calibri.
///
/// `Gift Budget and Tracker1.xlsx` (attached to #982) plots a stacked column
/// chart on a 1015.98 x 307.97pt frame whose value axis runs $0..$200 in $20
/// steps. Its own face is Segoe UI, which no runner is guaranteed to have; the
/// probe series below re-exported the same workbook with the chart's face
/// switched to Calibri, whose advances this crate holds natively, so every
/// figure is assertable wherever the tests run.
fn excel_gift_column_chart(size_pt: f64, number_format: Option<&str>) -> Chart {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Column;
    chart.host = crate::ir::ChartHost::Spreadsheet;
    chart.categories = vec!["Jan".to_string()];
    chart.series.truncate(1);
    // The overlaid line series reaches 180, which is what puts the automatic
    // axis at 200 in steps of 20 and so gives the $0..$200 labels.
    chart.series[0].values = vec![180.0];
    chart.has_legend = false;
    chart.title = None;
    chart.auto_title_deleted = true;
    chart.text_font_family = Some("Calibri".to_string());
    chart.value_axis_number_format = number_format.map(str::to_string);
    chart.value_axis_text_style.size_pt = Some(size_pt);
    chart
}

/// The frame `Gift Budget and Tracker1.xlsx` anchors its chart in, in the
/// chart's own points. The sheet prints at a 0.82 fit-to-page scale, so on the
/// page it measures 833.10 x 252.54pt.
const EXCEL_GIFT_CHART_FRAME: (f64, f64) = (1015.9784, 307.9732);
const EXCEL_GIFT_PRINT_SCALE: f64 = 0.82;
/// The frame's top edge in sheet points, where native Excel paints it: the
/// fitted export's 117.2703pt page seat undone by the 0.82 scale (#1542).
const EXCEL_GIFT_CHART_SPACE_FRAME_TOP: f64 = 143.012_536_585_365_85;
const EXCEL_GIFT_ANCHOR_Y: f64 = 79.256_692_913_385_83;

/// The same anchored worksheet combo chart with its three vertical-layout
/// text bands stated independently.
fn excel_gift_vertical_chart(
    value_axis_size_pt: f64,
    category_axis_size_pt: f64,
    legend_size_pt: f64,
) -> Chart {
    let mut chart = combo_budget_chart();
    chart.categories = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ]
    .into_iter()
    .map(str::to_string)
    .collect();
    chart.text_font_family = Some("Segoe UI".to_string());
    chart.value_axis_text_style.size_pt = Some(value_axis_size_pt);
    chart.category_axis_text_style.size_pt = Some(category_axis_size_pt);
    chart.legend_text_style.size_pt = Some(legend_size_pt);
    chart
}

fn anchored_excel_gift_chart_source(chart: Chart, chart_space_frame_top: f64) -> String {
    let margin_top: f64 = chart_space_frame_top * EXCEL_GIFT_PRINT_SCALE - EXCEL_GIFT_ANCHOR_Y;
    let doc: Document = make_doc(vec![Page::Sheet(SheetPage {
        name: "Gift budget and tracker".to_string(),
        size: PageSize::default(),
        margins: Margins {
            top: margin_top,
            ..Margins::default()
        },
        table: make_simple_table(vec![vec![""]]),
        header: None,
        footer: None,
        charts: vec![crate::ir::SheetChart {
            anchor_row: 3,
            placement: Some(crate::ir::SheetChartPlacement {
                x_offset_pt: 0.0,
                y_offset_pt: EXCEL_GIFT_ANCHOR_Y,
                width: EXCEL_GIFT_CHART_FRAME.0,
                height: EXCEL_GIFT_CHART_FRAME.1,
                print_scale: EXCEL_GIFT_PRINT_SCALE,
                clip_window: None,
            }),
            chart,
        }],
        images: Vec::new(),
        text_boxes: Vec::new(),
        shapes: Vec::new(),
    })]);
    generate_typst(&doc).unwrap().source
}

fn value_label_gridline_offset(source: &str, label: &str, tick_index: usize) -> f64 {
    let gridline_y: Vec<f64> = emitted_lines(source)
        .into_iter()
        .filter(|line| line.end_x > 100.0 && same_length(line.end_y, 0.0))
        .take(11)
        .map(|line| line.dy)
        .collect();
    assert_eq!(gridline_y.len(), 11, "expected eleven gridlines: {source}");
    let label_box: PlacedBox = placed_box_holding(source, label);
    label_box.dy + label_box.height / 2.0 - gridline_y[tick_index]
}

#[test]
fn an_excel_worksheet_column_plot_uses_the_native_vertical_edges() {
    // The native Excel export puts the chart-local plot at 11.000..259.180pt
    // inside the 307.9732pt frame, measured from the frame's own top edge in
    // sheet points; the unscaled control of #1607 reads the same 11pt inset
    // and 48.793pt bottom chrome. The old shared PowerPoint chrome produced a
    // plot 1.271pt short (issue #1250), and the #1250 calibration of
    // 10.146..258.327pt was taken against the converter's physical origin,
    // 0.854 sheet points below the frame Excel paints (issue #1607).
    let chart = excel_gift_vertical_chart(9.0, 9.0, 9.0);
    let (_, top, _, bottom) = axis_plot_rect(&chart, EXCEL_GIFT_CHART_FRAME, false);

    assert!(
        (top - 11.0).abs() <= 0.05,
        "plot top {top}pt, Excel's 11.000pt"
    );
    assert!(
        (bottom - 259.180).abs() <= 0.05,
        "plot bottom {bottom}pt, Excel's 259.180pt"
    );
    assert!(
        ((bottom - top) - 248.181).abs() <= 0.05,
        "plot height {}pt, Excel's 248.181pt",
        bottom - top
    );
}

#[test]
fn an_anchored_excel_worksheet_chart_snaps_interior_gridlines_in_sheet_space() {
    // Excel lays the #1471 chart out in unscaled sheet points, snaps each
    // interior major gridline to a whole point there, and only then applies
    // the sheet's 0.82 fit-to-page scale. The plot endpoints themselves stay
    // on the exact #1250 chrome model rather than being rounded.
    let mut chart = excel_gift_vertical_chart(9.0, 9.0, 9.0);
    chart.value_axis_max = Some(200.0);
    chart.value_axis_major_unit = Some(20.0);
    let source: String = anchored_excel_gift_chart_source(chart, EXCEL_GIFT_CHART_SPACE_FRAME_TOP);

    let page_frame_top = leading_pt(
        source
            .split_once("#place(top + left, dy: ")
            .expect("the sheet places its drawing layer chart")
            .1,
    )
    .expect("the chart page offset is a point measurement");
    let chart_space_frame_top = page_frame_top / EXCEL_GIFT_PRINT_SCALE;
    assert!(
        (chart_space_frame_top - EXCEL_GIFT_CHART_SPACE_FRAME_TOP).abs() <= 0.001,
        "test setup changed the sheet-space phase: {chart_space_frame_top}pt; source:\n{source}"
    );

    // Major gridlines are the first eleven full-width horizontal chart-chrome
    // lines: ticks are emitted from $0 at the bottom through $200 at the top.
    let gridline_y: Vec<f64> = emitted_lines(&source)
        .into_iter()
        .filter(|line| line.end_x > 100.0 && same_length(line.end_y, 0.0))
        .take(11)
        .map(|line| chart_space_frame_top + line.dy)
        .collect();
    assert_eq!(
        gridline_y.len(),
        11,
        "expected eleven major gridlines: {source}"
    );

    let expected = [
        402.193_536_585_365_85,
        377.0,
        353.0,
        328.0,
        303.0,
        278.0,
        253.0,
        228.0,
        204.0,
        179.0,
        154.012_536_585_365_85,
    ];
    for (index, (actual, expected)) in gridline_y.into_iter().zip(expected).enumerate() {
        assert!(
            (actual - expected).abs() <= 0.002,
            "gridline {index} lands at {actual} sheet pt, native Excel uses {expected} sheet pt; source:\n{source}"
        );
    }
}

#[test]
fn an_anchored_excel_worksheet_chart_snaps_value_labels_independently() {
    // Native Excel rounds the value-label baseline separately from the
    // gridline it accompanies. On the #1499 chart that gives the $140 label
    // one extra unscaled sheet point while the adjacent $120 and $160 labels
    // keep the common gridline-relative seat.
    let mut chart: Chart = excel_gift_vertical_chart(9.0, 9.0, 9.0);
    chart.value_axis_max = Some(200.0);
    chart.value_axis_major_unit = Some(20.0);
    chart.value_axis_number_format = Some("\"$\"#,##0".to_string());
    let source: String = anchored_excel_gift_chart_source(chart, EXCEL_GIFT_CHART_SPACE_FRAME_TOP);
    let offsets: [f64; 3] = [
        value_label_gridline_offset(&source, "\\$120", 6),
        value_label_gridline_offset(&source, "\\$140", 7),
        value_label_gridline_offset(&source, "\\$160", 8),
    ];
    let expected: [f64; 3] = [0.0, 1.0, 0.0];
    for ((label, actual), expected) in ["$120", "$140", "$160"]
        .into_iter()
        .zip(offsets)
        .zip(expected)
    {
        assert!(
            (actual - expected).abs() <= 0.002,
            "{label} label offset is {actual} sheet pt, native Excel uses {expected} sheet pt; source:\n{source}"
        );
    }
}

#[test]
fn excel_value_label_snap_follows_the_sheet_phase_instead_of_the_tick_value() {
    // Moving the whole chart 0.2 sheet points changes which continuous label
    // coordinate crosses the next integer boundary. The extra point therefore
    // moves from $140 to $120; it is not a special case for one tick value.
    let mut chart: Chart = excel_gift_vertical_chart(9.0, 9.0, 9.0);
    chart.value_axis_max = Some(200.0);
    chart.value_axis_major_unit = Some(20.0);
    chart.value_axis_number_format = Some("\"$\"#,##0".to_string());
    let source: String =
        anchored_excel_gift_chart_source(chart, EXCEL_GIFT_CHART_SPACE_FRAME_TOP + 0.2);
    let offsets: [f64; 3] = [
        value_label_gridline_offset(&source, "\\$120", 6),
        value_label_gridline_offset(&source, "\\$140", 7),
        value_label_gridline_offset(&source, "\\$160", 8),
    ];
    let expected: [f64; 3] = [1.0, 0.0, 0.0];
    for ((label, actual), expected) in ["$120", "$140", "$160"]
        .into_iter()
        .zip(offsets)
        .zip(expected)
    {
        assert!(
            (actual - expected).abs() <= 0.002,
            "shifted {label} label offset is {actual} sheet pt, expected {expected} sheet pt; source:\n{source}"
        );
    }
}

/// Every column edge Excel for Mac 16.112 paints on the public gift workbook
/// lands on a whole unscaled sheet point. Re-zip-controlled native exports
/// through `scripts/probes/issue-1543-column-gap-width.json` and
/// `scripts/probes/issue-1543-column-overlap.json` paint 50 columns whose
/// every edge but one is `round(continuous edge)`; the one exception is the
/// gapWidth-0 first column's left edge, which stops at the plot's exact left
/// edge instead of the rounded point beyond it (#1543).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn gift_budget_compiled_columns_land_on_native_whole_sheet_points() {
    let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1603_gift_budget.xlsx");
    /// One native probe export: its band factors and the sheet-point edges of
    /// every painted column, left to right, with `true` marking a left edge
    /// that native clips to the plot rather than rounds.
    struct NativeColumnProbe {
        gap_width_percent: f64,
        overlap_percent: f64,
        native_edges: &'static [(f64, f64, bool)],
    }
    let probe = |gap_width_percent: f64,
                 overlap_percent: f64,
                 native_edges: &'static [(f64, f64, bool)]| NativeColumnProbe {
        gap_width_percent,
        overlap_percent,
        native_edges,
    };
    let cases: [NativeColumnProbe; 7] = [
        probe(
            150.0,
            100.0,
            &[
                (412.0, 445.0, false),
                (655.0, 687.0, false),
                (816.0, 849.0, false),
                (816.0, 849.0, false),
                (897.0, 930.0, false),
            ],
        ),
        probe(
            0.0,
            100.0,
            &[
                (388.109, 469.0, true),
                (631.0, 711.0, false),
                (792.0, 873.0, false),
                (792.0, 873.0, false),
                (873.0, 954.0, false),
            ],
        ),
        probe(
            50.0,
            100.0,
            &[
                (402.0, 455.0, false),
                (644.0, 698.0, false),
                (806.0, 860.0, false),
                (806.0, 860.0, false),
                (887.0, 940.0, false),
            ],
        ),
        probe(
            219.0,
            100.0,
            &[
                (416.0, 441.0, false),
                (658.0, 684.0, false),
                (820.0, 845.0, false),
                (820.0, 845.0, false),
                (901.0, 926.0, false),
            ],
        ),
        probe(
            500.0,
            100.0,
            &[
                (422.0, 435.0, false),
                (664.0, 678.0, false),
                (826.0, 839.0, false),
                (826.0, 839.0, false),
                (907.0, 920.0, false),
            ],
        ),
        probe(
            150.0,
            0.0,
            &[
                (402.0, 420.0, false),
                (680.0, 698.0, false),
                (806.0, 824.0, false),
                (842.0, 860.0, false),
                (904.0, 922.0, false),
            ],
        ),
        probe(
            150.0,
            -27.0,
            &[
                (400.0, 416.0, false),
                (683.0, 699.0, false),
                (804.0, 820.0, false),
                (845.0, 861.0, false),
                (905.0, 921.0, false),
            ],
        ),
    ];
    let series_paints = [
        typst::visualize::Color::from_u8(0x3D, 0x2B, 0x2D, 255),
        typst::visualize::Color::from_u8(0x7A, 0x63, 0x69, 255),
        typst::visualize::Color::from_u8(0x5D, 0x37, 0x54, 255),
    ];
    for NativeColumnProbe {
        gap_width_percent,
        overlap_percent,
        native_edges,
    } in cases
    {
        let (mut doc, _) = crate::parser::Parser::parse(
            &crate::parser::xlsx::XlsxParser,
            data,
            &crate::config::ConvertOptions::default(),
        )
        .expect("the public gift-budget fixture parses");
        let mut patched: usize = 0;
        for page in &mut doc.pages {
            if let crate::ir::Page::Sheet(sheet) = page {
                for anchored in &mut sheet.charts {
                    anchored.chart.bar_band_layout.gap_width_percent = gap_width_percent;
                    anchored.chart.bar_band_layout.overlap_percent = overlap_percent;
                    patched += 1;
                }
            }
        }
        assert_eq!(patched, 1, "the workbook anchors one chart");
        let output = generate_typst(&doc).unwrap();
        let pages =
            crate::render::pdf::compiled_page_paint_sequences(&output.source, &output.images)
                .expect("the gift-budget column variant compiles");
        assert_eq!(pages.len(), 2);
        let mut columns: Vec<(f64, f64)> = pages[1]
            .iter()
            .filter(|paint| {
                let (left, top, right, bottom) = paint.bounds;
                paint
                    .rectangle_fill
                    .as_ref()
                    .is_some_and(|fill| series_paints.contains(fill))
                    && left > 300.0
                    && right < 1125.0
                    && top > 120.0
                    && bottom < 335.0
                    && bottom - top > 1.0
            })
            .map(|paint| (paint.bounds.0, paint.bounds.2))
            .collect();
        columns.sort_by(|a, b| a.0.total_cmp(&b.0).then(a.1.total_cmp(&b.1)));
        assert_eq!(
            columns.len(),
            native_edges.len(),
            "gapWidth {gap_width_percent} overlap {overlap_percent}: painted columns {columns:?}"
        );
        for (index, ((left, right), (native_left, native_right, clipped_to_plot))) in
            columns.iter().zip(native_edges).enumerate()
        {
            let expected_left: f64 = native_left * EXCEL_GIFT_PRINT_SCALE;
            let expected_right: f64 = native_right * EXCEL_GIFT_PRINT_SCALE;
            // A rounded edge reproduces exactly; the plot-clipped edge follows
            // our own plot edge, which is held to the 0.5pt Excel gate.
            let tolerance: f64 = if *clipped_to_plot { 0.5 } else { 0.01 };
            assert!(
                (left - expected_left).abs() <= tolerance && (right - expected_right).abs() <= 0.01,
                "gapWidth {gap_width_percent} overlap {overlap_percent} column {index}: \
                 native {expected_left:.4}..{expected_right:.4}pt, got {left:.4}..{right:.4}pt"
            );
        }
    }
}

/// The whole-point rule is a property of anchored worksheet columns, not of
/// the gift workbook's data: at every gap width the generator's column widths
/// and the spacing between column starts are whole sheet points, while a
/// column the rounding would push past the plot stops at the plot edge.
#[test]
fn an_anchored_excel_worksheet_chart_rounds_column_edges_to_whole_sheet_points() {
    for gap_width_percent in [0.0, 50.0, 100.0, 150.0, 219.0, 300.0, 500.0] {
        let mut chart = excel_gift_vertical_chart(9.0, 9.0, 9.0);
        chart.bar_band_layout.gap_width_percent = gap_width_percent;
        chart.bar_band_layout.overlap_percent = 100.0;
        let (plot_left, _, plot_right, _) = axis_plot_rect(&chart, EXCEL_GIFT_CHART_FRAME, false);
        let source: String =
            anchored_excel_gift_chart_source(chart, EXCEL_GIFT_CHART_SPACE_FRAME_TOP);
        // The plot's rectangles, one per category and series, are the only
        // rectangles above the legend band.
        let columns: Vec<PlacedRect> = emitted_rects(&source)
            .into_iter()
            .filter(|rect| rect.dy < 270.0)
            .collect();
        assert_eq!(
            columns.len(),
            24,
            "gapWidth {gap_width_percent}: twelve categories of two column series; got:\n{source}"
        );
        let whole = |value: f64| (value - value.round()).abs() <= 1e-6;
        let first: PlacedRect = columns[0];
        let last: PlacedRect = columns[columns.len() - 1];
        if gap_width_percent == 0.0 {
            assert!(
                (first.dx - plot_left).abs() <= 1e-6
                    && (last.dx + last.width - plot_right).abs() <= 1e-6,
                "gapWidth 0: the outer columns stop at the plot edges {plot_left}..{plot_right}; \
                 got {first:?} and {last:?} in:\n{source}"
            );
        }
        // The second category's column is interior at every gap width, so its
        // rounded left edge is the reference the other edges must be whole
        // sheet points away from.
        let reference: f64 = columns[2].dx;
        for (index, rect) in columns.iter().enumerate() {
            let is_outer: bool = gap_width_percent == 0.0 && !(2..22).contains(&index);
            if is_outer {
                continue;
            }
            assert!(
                whole(rect.width) && whole(rect.dx - reference),
                "gapWidth {gap_width_percent} column {index}: {rect:?} is not on whole sheet points \
                 relative to the second category's {reference}pt; got:\n{source}"
            );
        }
    }
}

#[test]
fn an_excel_worksheet_plot_top_follows_the_native_value_size_probes() {
    // Re-zip-controlled native Excel exports with only c:valAx text size
    // changed. Excel floors the top chrome at the 9pt seat, then grows it by
    // approximately two thirds of a point per additional text point.
    let measurements = [
        (7.0, 11.0000),
        (9.0, 11.0000),
        (11.0, 12.3150),
        (14.0, 14.3100),
        (18.0, 16.9726),
    ];
    for (size_pt, expected_top) in measurements {
        let chart = excel_gift_vertical_chart(size_pt, 9.0, 9.0);
        let actual_top = axis_plot_rect(&chart, EXCEL_GIFT_CHART_FRAME, false).1;
        assert!(
            (actual_top - expected_top).abs() <= 0.05,
            "{size_pt}pt value axis: plot top {actual_top}pt, Excel's {expected_top}pt"
        );
    }
}

#[test]
fn an_excel_worksheet_plot_bottom_follows_both_native_text_band_probes() {
    // The bottom edge moves independently with the category and bottom-legend
    // bands: 2.05pt per category text point and 1.3167pt per legend text point.
    let category_measurements = [
        (7.0, 263.2764),
        (9.0, 259.1800),
        (11.0, 255.0800),
        (14.0, 248.9332),
        (18.0, 240.7282),
    ];
    for (size_pt, expected_bottom) in category_measurements {
        let chart = excel_gift_vertical_chart(9.0, size_pt, 9.0);
        let actual_bottom = axis_plot_rect(&chart, EXCEL_GIFT_CHART_FRAME, false).3;
        assert!(
            (actual_bottom - expected_bottom).abs() <= 0.05,
            "{size_pt}pt category axis: plot bottom {actual_bottom}pt, Excel's {expected_bottom}pt"
        );
    }

    let legend_measurements = [
        (7.0, 261.8134),
        (9.0, 259.1800),
        (11.0, 256.5467),
        (14.0, 252.5964),
        (18.0, 247.3247),
    ];
    for (size_pt, expected_bottom) in legend_measurements {
        let chart = excel_gift_vertical_chart(9.0, 9.0, size_pt);
        let actual_bottom = axis_plot_rect(&chart, EXCEL_GIFT_CHART_FRAME, false).3;
        assert!(
            (actual_bottom - expected_bottom).abs() <= 0.05,
            "{size_pt}pt legend: plot bottom {actual_bottom}pt, Excel's {expected_bottom}pt"
        );
    }
}

#[test]
fn an_excel_worksheet_plot_without_a_legend_reclaims_excels_legend_band() {
    // Removing c:legend in the native workbook extends the plot by 23.850pt.
    let mut chart = excel_gift_vertical_chart(9.0, 9.0, 9.0);
    chart.has_legend = false;
    let actual_bottom = axis_plot_rect(&chart, EXCEL_GIFT_CHART_FRAME, false).3;
    assert!(
        (actual_bottom - 283.030).abs() <= 0.05,
        "plot bottom without legend {actual_bottom}pt, Excel's 283.030pt"
    );
}

#[test]
fn an_excel_worksheet_chart_seats_its_category_and_bottom_legend_bands() {
    // A fresh Excel for Mac 16.112 export of `Gift Budget and Tracker1.xlsx`
    // (SHA-256 25f5dc75dab19ea12042979a61842314ddc226e3e45d447e36b2a2a104112613)
    // puts every category-label baseline at 343.580pt and every bottom-legend
    // baseline at 359.980pt on the printed page. The pre-fix renderer put the
    // same baselines at 345.786pt and 364.196pt. Undoing the sheet's exact
    // 0.82 print scale makes those residuals 2.690pt and 5.141pt in chart-local
    // coordinates (#1240). Those seats were taken against the converter's
    // physical frame origin; the chart-local tops below are the same native
    // baselines measured from the frame Excel paints, 0.853 sheet points
    // higher (#1607).
    //
    // Restate the combo chart in Calibri because its native metrics are held
    // by the crate and therefore do not depend on a CI runner having Segoe UI.
    // For this fixed two-category chart the labels remain flat, so series and
    // category counts do not enter the measured seats; the anchored frame and
    // the three 9pt text bands are the controlled factors. Larger category
    // sets can rotate labels and take the separate rotated-label layout path.
    let mut chart = combo_budget_chart();
    chart.text_font_family = Some("Calibri".to_string());
    chart.category_axis_text_style.size_pt = Some(9.0);
    chart.legend_text_style.size_pt = Some(9.0);
    chart.value_axis_text_style.size_pt = Some(9.0);
    let source = framed_chart_source(&chart, EXCEL_GIFT_CHART_FRAME.0, EXCEL_GIFT_CHART_FRAME.1);

    let category_top = placed_box_holding(&source, "May").dy;
    let legend_top = legend_entry_y(&source, "Birthday Budget");

    assert!(
        (category_top - 257.536).abs() <= 0.02,
        "category-label band top {category_top}pt, native-derived 257.536pt; got:\n{source}"
    );
    assert!(
        (legend_top - 289.685).abs() <= 0.02,
        "bottom-legend band top {legend_top}pt, native-derived 289.685pt; got:\n{source}"
    );
}

#[test]
fn excel_worksheet_band_seats_follow_the_native_size_probes() {
    // Chart-local box tops that reproduce the five native baselines recorded
    // by the one-factor probes. Each series changes one text size and leaves
    // the frame, plot data and other two chart text bands at 9pt.
    let category_tops = [
        (7.0, 261.6362),
        (9.0, 257.5362),
        (11.0, 252.4362),
        (14.0, 245.2862),
        (18.0, 236.0862),
    ];
    for (size_pt, expected_top) in category_tops {
        let mut chart = combo_budget_chart();
        chart.text_font_family = Some("Calibri".to_string());
        chart.category_axis_text_style.size_pt = Some(size_pt);
        chart.legend_text_style.size_pt = Some(9.0);
        chart.value_axis_text_style.size_pt = Some(9.0);
        let source =
            framed_chart_source(&chart, EXCEL_GIFT_CHART_FRAME.0, EXCEL_GIFT_CHART_FRAME.1);
        let actual_top = placed_box_holding(&source, "May").dy;
        assert!(
            (actual_top - expected_top).abs() <= 0.001,
            "{size_pt}pt category top {actual_top}pt, expected {expected_top}pt"
        );
    }

    let legend_tops = [
        (7.0, 292.0852),
        (9.0, 289.6852),
        (11.0, 287.2852),
        (14.0, 284.1852),
        (18.0, 280.3852),
    ];
    for (size_pt, expected_top) in legend_tops {
        let mut chart = combo_budget_chart();
        chart.text_font_family = Some("Calibri".to_string());
        chart.category_axis_text_style.size_pt = Some(9.0);
        chart.legend_text_style.size_pt = Some(size_pt);
        chart.value_axis_text_style.size_pt = Some(9.0);
        let source =
            framed_chart_source(&chart, EXCEL_GIFT_CHART_FRAME.0, EXCEL_GIFT_CHART_FRAME.1);
        let actual_top = legend_entry_y(&source, "Birthday Budget");
        assert!(
            (actual_top - expected_top).abs() <= 0.001,
            "{size_pt}pt legend top {actual_top}pt, expected {expected_top}pt"
        );
    }
}

#[test]
fn non_worksheet_chart_hosts_keep_their_category_and_legend_seats() {
    for host in [
        crate::ir::ChartHost::Presentation,
        crate::ir::ChartHost::SpreadsheetChartsheet,
        crate::ir::ChartHost::WordProcessing,
    ] {
        let mut chart = combo_budget_chart();
        chart.host = host;
        chart.text_font_family = Some("Calibri".to_string());
        chart.category_axis_text_style.size_pt = Some(9.0);
        chart.legend_text_style.size_pt = Some(9.0);
        chart.value_axis_text_style.size_pt = Some(9.0);
        let source =
            framed_chart_source(&chart, EXCEL_GIFT_CHART_FRAME.0, EXCEL_GIFT_CHART_FRAME.1);
        let plot_bottom = axis_plot_rect(&chart, EXCEL_GIFT_CHART_FRAME, false).3;
        let category_top = placed_box_holding(&source, "May").dy;
        let legend_top = legend_entry_y(&source, "Birthday Budget");

        assert!(
            (category_top - (plot_bottom + 2.0)).abs() <= 0.001,
            "{host:?} category seat changed: {category_top}pt after {plot_bottom}pt"
        );
        assert!(
            (legend_top - (EXCEL_GIFT_CHART_FRAME.1 - 14.0)).abs() <= 0.001,
            "{host:?} bottom-legend seat changed: {legend_top}pt"
        );
    }
}

#[test]
fn a_framed_excel_column_chart_seats_its_plot_where_excel_measures_it() {
    // Native Excel for Mac 16.112 exports of the #1166 workbook, one factor
    // changed per variant, measured off `mutool draw -F trace` as the value
    // gridlines' left end relative to the chart frame. Excel seats the widest
    // tick label 6.5pt inside the frame whatever the size or the label, then
    // leaves 5/6 of the face's ascent plus half its descent before the plot.
    //
    // The size series and the number-format series both have to be here: a
    // model fitted to one alone can trade the label's own width against the
    // clearance after it and still reproduce that series exactly.
    let measurements = [
        (7.0, Some("\\$#,##0"), 27.1941),
        (9.0, Some("\\$#,##0"), 33.0891),
        (11.0, Some("\\$#,##0"), 39.0068),
        (14.0, Some("\\$#,##0"), 47.8684),
        (18.0, Some("\\$#,##0"), 59.7010),
        (9.0, None, 28.5291),
        (9.0, Some("\\$#,##0.0000"), 53.5991),
    ];

    for (size_pt, number_format, expected_left) in measurements {
        let chart = excel_gift_column_chart(size_pt, number_format);
        let actual_left = axis_plot_rect(&chart, EXCEL_GIFT_CHART_FRAME, false).0;
        assert!(
            (actual_left - expected_left).abs() <= 0.05,
            "{size_pt}pt {number_format:?}: plot left {actual_left}pt, Excel's {expected_left}pt"
        );
    }
}

#[test]
fn an_excel_column_value_label_ends_where_excel_ends_it() {
    // The same exports place every widest label's first glyph 6.5pt inside the
    // frame. The labels are right-aligned in their box, so the box has to end
    // where the clearance before the plot begins; sizing it independently of
    // the gutter is what left the #1166 labels 2.63pt short of Excel's.
    let chart = excel_gift_column_chart(9.0, Some("\\$#,##0"));
    let source = framed_chart_source(&chart, EXCEL_GIFT_CHART_FRAME.0, EXCEL_GIFT_CHART_FRAME.1);
    let widest = source
        .lines()
        .find(|line| line.contains("align(right + horizon)") && line.ends_with(r"[\$100]]])"))
        .expect("the $100 value-axis label is emitted");
    let dx: f64 = leading_pt(widest.split_once("dx: ").expect("the label is placed").1)
        .expect("the label's dx is a measurement");
    let box_w: f64 = leading_pt(
        widest
            .split_once("box(width: ")
            .expect("the label sits in a box")
            .1,
    )
    .expect("the label box's width is a measurement");
    // Calibri advances `$100` at 2.027344em, so 6.5pt of frame inset plus the
    // label is 24.746pt — where Excel's own widest label starts and ends.
    assert!(
        (dx + box_w - 24.746).abs() <= 0.01,
        "the value-label box ends at {}pt, Excel's at 24.746pt: {widest}",
        dx + box_w
    );
    assert!(
        (dx - 0.0).abs() < 1e-9,
        "the spreadsheet label box keeps its origin: {widest}"
    );
}

#[test]
fn a_non_powerpoint_or_unsized_column_keeps_its_existing_label_origin() {
    let mut spreadsheet = crowded_column_chart();
    spreadsheet.host = crate::ir::ChartHost::Spreadsheet;
    spreadsheet.text_style.size_pt = Some(11.97);
    let spreadsheet_source = framed_chart_source(&spreadsheet, 401.95, 344.25);
    let spreadsheet_zero = spreadsheet_source
        .lines()
        .find(|line| line.contains("align(right + horizon)") && line.ends_with("[0]]])"))
        .expect("the spreadsheet zero value-axis label is emitted");
    assert!(spreadsheet_zero.contains("dx: 0pt"), "{spreadsheet_zero}");

    let mut unsized_powerpoint = crowded_column_chart();
    unsized_powerpoint.host = crate::ir::ChartHost::Presentation;
    let powerpoint_source = framed_chart_source(&unsized_powerpoint, 401.95, 344.25);
    let powerpoint_zero = powerpoint_source
        .lines()
        .find(|line| line.contains("align(right + horizon)") && line.ends_with("[0]]])"))
        .expect("the unsized PowerPoint zero value-axis label is emitted");
    assert!(powerpoint_zero.contains("dx: 0pt"), "{powerpoint_zero}");
}

#[test]
fn a_chart_title_occupies_the_same_fixed_band_used_by_plot_geometry() {
    let mut chart = bar_chart_at(Some(18.0), &["Q1", "Q2"]);
    chart.title = Some("Sales".to_string());
    let source = framed_chart_source(&chart, 480.0, 320.0);
    assert!(
        source.contains("#block(width: 480pt, height: 46.21pt, above: 0pt, below: 0pt)"),
        "the emitted title must occupy its measured plot band and frame width, got:\n{source}"
    );
    assert!(
        !source.contains("#block(width: 100%, height: 46.21pt"),
        "a framed title must not resolve 100% against the slide, got:\n{source}"
    );
}

/// PowerPoint measures a title's edge-mode x/y against the full chart frame.
/// Its text begins 3pt inside that title box, and the title still reserves the
/// existing plot band. `c:title/c:overlay` is not modelled yet, so this test
/// covers manual position without changing that reservation (issue #1423).
#[test]
fn a_powerpoint_chart_title_uses_its_manual_edge_anchor() {
    let mut chart = bar_chart_at(Some(18.62), &["Q1", "Q2"]);
    chart.host = crate::ir::ChartHost::Presentation;
    chart.title = Some("Annual Income & Gross Profit".to_string());
    chart.title_text_style.size_pt = Some(18.62);
    chart.title_layout = Some(crate::ir::ChartTitleLayout {
        x: 0.007212769679107099,
        y: 0.0001182572865366946,
    });

    let source = framed_chart_source(&chart, 689.187, 221.058);

    assert!(
        source.contains(
            "#place(top + left, dx: 7.970947096834784pt, dy: 0.026141719247228634pt, text(top-edge: 18.62pt, bottom-edge: \"baseline\", size: 18.62pt"
        ),
        "the title must use its chart-relative edge anchor and text inset, got:\n{source}"
    );
    assert!(
        source.contains("#block(width: 689.187pt, height: 41.1902144pt"),
        "the manually positioned title must keep the existing title band, got:\n{source}"
    );
}

#[test]
fn framed_line_radar_and_pie_charts_center_their_titles_in_the_frame() {
    let mut line = two_series_bar_chart(Vec::new());
    line.chart_type = ChartType::Line;
    line.title = Some("Line title".to_string());
    line.categories = vec!["Q1".to_string(), "Q2".to_string()];
    line.series[0].values = vec![1.0, 2.0];
    line.series[1].values = vec![2.0, 1.0];

    let mut radar = radar_chart();
    radar.title = Some("Radar title".to_string());

    let mut pie = pie_chart(vec![60.0, 40.0]);
    pie.title = Some("Pie title".to_string());

    for (name, chart) in [("line", line), ("radar", radar), ("pie", pie)] {
        let source = framed_chart_source(&chart, 321.0, 240.0);
        assert!(
            source.contains("#block(width: 321pt)[#align(center)"),
            "the {name} title must use its chart frame, got:\n{source}"
        );
    }
}

/// `c:chartSpace/c:spPr` belongs to the whole chart space, not only to the
/// plot left after the title takes its band. Every plot family must therefore
/// open the full framed outline before it writes the title (issue #1216).
#[test]
fn every_plot_family_draws_a_titled_chart_inside_its_full_area_outline() {
    let mut axis = bar_chart_at(Some(14.0), &["Q1", "Q2"]);
    axis.title = Some("Axis title".to_string());

    let mut line = two_series_bar_chart(Vec::new());
    line.chart_type = ChartType::Line;
    line.title = Some("Line title".to_string());
    line.categories = vec!["Q1".to_string(), "Q2".to_string()];
    line.series[0].values = vec![1.0, 2.0];
    line.series[1].values = vec![2.0, 1.0];

    let mut radar = radar_chart();
    radar.title = Some("Radar title".to_string());

    let mut pie = pie_chart(vec![60.0, 40.0]);
    pie.title = Some("Pie title".to_string());

    let outline = ChartAreaOutline::Explicit {
        width_pt: Some(2.0),
        color: Some(crate::ir::Color::new(0xd9, 0xd9, 0xd9)),
    };
    let area_start =
        "#box(width: 321pt, height: 240pt, fill: none, stroke: 2pt + rgb(217, 217, 217))[";

    for (family, title, mut chart) in [
        ("axis", "Axis title", axis),
        ("line", "Line title", line),
        ("radar", "Radar title", radar),
        ("pie", "Pie title", pie),
    ] {
        chart.chart_area_outline = outline.clone();
        let source = framed_chart_source(&chart, 321.0, 240.0);
        let area_position = source.find(area_start).unwrap_or_else(|| {
            panic!("the {family} chart must outline its full frame, got:\n{source}")
        });
        let title_position = source
            .find(title)
            .unwrap_or_else(|| panic!("the {family} chart must print its title: {source}"));
        assert!(
            area_position < title_position,
            "the {family} title must sit inside the chart-area outline, got:\n{source}"
        );
        assert_eq!(
            source.matches("stroke: 2pt + rgb(217, 217, 217))[").count(),
            1,
            "the {family} chart must draw exactly one area outline, got:\n{source}"
        );
    }
}

#[test]
fn a_flowed_chart_title_keeps_its_container_width() {
    let mut chart = bar_chart_at(Some(18.0), &["Q1", "Q2"]);
    chart.title = Some("Sales".to_string());
    let source = chart_source(chart);
    assert!(source.contains("#block(width: 100%, height: 46.21pt, above: 0pt, below: 0pt)"));
}

#[test]
fn the_category_gutter_never_narrows_below_the_calibrated_width() {
    // The gutter is measured from the widest label, and a face that cannot be
    // measured — wasm has no font search — must not collapse it.
    let mut chart = bar_chart_at(Some(18.0), &["Q", "R"]);
    chart.text_font_family = Some("Definitely Missing Chart Face 706".to_string());
    assert_eq!(chart_category_gutter_pt(&chart), LABEL_W + GAP);
}

#[test]
fn the_category_gutter_grows_with_the_widest_label() {
    // A width holding text has to follow what the text says, not just its size.
    // `bar-chart.pptx`'s labels are as short as `4th Qtr`, so scaling the flat
    // constant by the band's 1.8 would reserve far more than they need.
    let short: f64 = chart_category_gutter_pt(&bar_chart_at(Some(18.0), &["Q1", "Q2"]));
    let long: f64 = chart_category_gutter_pt(&bar_chart_at(
        Some(18.0),
        &["Q1", "A considerably longer category label"],
    ));
    assert!(
        long > short,
        "a longer label must widen the gutter: {long} against {short}"
    );
}

#[test]
fn horizontal_category_labels_stop_the_text_scaled_clearance_short_of_the_plot() {
    let chart = bar_chart_at(Some(18.0), &["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"]);
    let source = framed_chart_source(&chart, 480.0, 320.0);
    let label_width = chart_category_gutter_pt(&chart) - 16.686;
    let label = source
        .lines()
        .find(|line| line.contains("[4th Qtr]"))
        .expect("the chart emits its category label");
    assert!(
        label.contains(&format!("box(width: {}pt", format_f64(label_width))),
        "an 18pt label must stop 16.686pt short of the plot, got:\n{label}"
    );
}

#[test]
fn horizontal_category_labels_keep_the_legacy_gap_without_a_declared_size() {
    let chart = bar_chart_at(None, &["Q1", "Q2"]);
    let source = framed_chart_source(&chart, 480.0, 320.0);
    let label = source
        .lines()
        .find(|line| line.contains("[Q2]"))
        .expect("the chart emits its category label");
    assert!(
        label.contains(&format!("box(width: {}pt", format_f64(LABEL_W))),
        "an undeclared-size label must keep the pre-#706 6pt gap, got:\n{label}"
    );
}

#[test]
fn horizontal_category_labels_keep_the_fallback_width_when_the_face_is_unmeasurable() {
    let mut chart = bar_chart_at(Some(18.0), &["Q1", "Q2"]);
    chart.text_font_family = Some("Definitely Missing Chart Face 998".to_string());
    let source = framed_chart_source(&chart, 480.0, 320.0);
    let label = source
        .lines()
        .find(|line| line.contains("[Q2]"))
        .expect("the chart emits its category label");
    assert!(
        label.contains(&format!("box(width: {}pt", format_f64(LABEL_W))),
        "an unmeasurable face must keep the calibrated 62pt fallback, got:\n{label}"
    );
}

// ----- The automatic chart-area outline is host-dependent (issue #823) -----

/// The `#box(...)` line that opens the chart area, whose `stroke:` is the
/// chart-area outline. The gridlines below it repeat the same stroke string.
fn chart_area_box_line(source: &str) -> &str {
    source
        .lines()
        .find(|line| line.starts_with("#box(width:") && line.contains("stroke:"))
        .expect("the chart opens an area box")
}

fn framed_bar_chart_on(host: crate::ir::ChartHost) -> Chart {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Bar;
    chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
    chart.series[0].values = vec![4.0, 8.0];
    chart.series[1].values = vec![2.0, 6.0];
    chart.host = host;
    chart
}

#[test]
fn a_slide_chart_draws_no_automatic_area_outline() {
    // PowerPoint draws none. Applying Excel's default everywhere put a
    // 480 x 301pt rectangle around every chart on a slide.
    let source: String = chart_source(framed_bar_chart_on(crate::ir::ChartHost::Presentation));
    // The gridlines carry the same stroke string legitimately, so the box's own
    // `stroke:` is what has to be read, not the source as a whole.
    let box_line: &str = chart_area_box_line(&source);
    assert!(
        box_line.contains("stroke: none"),
        "a slide chart must not draw Excel's automatic border, got: {box_line}"
    );
}

#[test]
fn a_workbook_chart_keeps_the_measured_excel_outline() {
    // #637 measured this against a native Excel export and it must not move.
    for host in [
        crate::ir::ChartHost::Spreadsheet,
        crate::ir::ChartHost::SpreadsheetChartsheet,
    ] {
        let source: String = chart_source(framed_bar_chart_on(host));
        let box_line: &str = chart_area_box_line(&source);
        assert!(
            box_line.contains(CHART_AREA_OUTLINE),
            "a workbook chart keeps Excel's automatic border on {host:?}, got: {box_line}"
        );
    }
}

#[test]
fn an_explicit_outline_survives_on_every_host() {
    // Only the *automatic* default is host-dependent; a chart that states a
    // line gets it wherever it lives, and `noFill` still suppresses.
    for host in [
        crate::ir::ChartHost::Presentation,
        crate::ir::ChartHost::Spreadsheet,
        crate::ir::ChartHost::SpreadsheetChartsheet,
        crate::ir::ChartHost::WordProcessing,
    ] {
        let mut chart = framed_bar_chart_on(host);
        chart.chart_area_outline = ChartAreaOutline::Explicit {
            width_pt: Some(2.0),
            color: Some(crate::ir::Color::new(0xd9, 0xd9, 0xd9)),
        };
        let source: String = chart_source(chart);
        assert!(
            chart_area_box_line(&source).contains("2pt + rgb(217, 217, 217)"),
            "an explicit outline must survive on {host:?}"
        );

        let mut chart = framed_bar_chart_on(host);
        chart.chart_area_outline = ChartAreaOutline::Suppressed;
        let source: String = chart_source(chart);
        assert!(
            chart_area_box_line(&source).contains("stroke: none"),
            "on {host:?}"
        );
    }
}

// ----- The automatic horizontal value-axis scale is host-dependent (#824) -----

fn auto_scaled_bar_chart_on(host: crate::ir::ChartHost) -> Chart {
    let mut chart = framed_bar_chart_on(host);
    chart.text_style.size_pt = Some(18.0);
    chart.series[0].values = vec![8.2, 3.2];
    chart.series[1].values = vec![1.4, 1.2];
    chart
}

#[test]
fn a_slide_bar_chart_uses_powerpoints_measured_auto_scale() {
    let source = chart_source(auto_scaled_bar_chart_on(crate::ir::ChartHost::Presentation));
    assert_eq!(
        emitted_axis_ticks_at_size(&source, 18.0),
        vec![0.0, 2.0, 4.0, 6.0, 8.0, 10.0]
    );
}

#[test]
fn a_workbook_bar_chart_keeps_excels_measured_auto_scale() {
    let source = chart_source(auto_scaled_bar_chart_on(crate::ir::ChartHost::Spreadsheet));
    assert_eq!(
        emitted_axis_ticks_at_size(&source, 18.0),
        vec![0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0]
    );
}

// ----- A horizontal legend advances by each entry's width (issue #827) -----

fn legend_entry_x(source: &str, label: &str) -> f64 {
    let marker: String = format!("[{label}]])");
    let index: usize = source.find(&marker).expect("the entry is drawn");
    let line_start: usize = source[..index].rfind('\n').map_or(0, |at| at + 1);
    let line: &str = &source[line_start..index];
    line.split("dx: ")
        .nth(1)
        .and_then(|rest| rest.split("pt").next())
        .and_then(|value| value.trim().parse::<f64>().ok())
        .expect("the entry is placed")
}

fn legend_entry_y(source: &str, label: &str) -> f64 {
    let marker: String = format!("[{label}]])");
    let index: usize = source.find(&marker).expect("the entry is drawn");
    let line_start: usize = source[..index].rfind('\n').map_or(0, |at| at + 1);
    let line: &str = &source[line_start..index];
    line.split("dy: ")
        .nth(1)
        .and_then(|rest| rest.split("pt").next())
        .and_then(|value| value.trim().parse::<f64>().ok())
        .expect("the entry is placed")
}

fn horizontal_value_axis_label_y(source: &str, label: &str) -> f64 {
    let marker: String = format!(")[{label}]]])");
    let line = source
        .lines()
        .find(|line| line.contains("box(width: 24pt)") && line.contains(&marker))
        .expect("the horizontal value-axis label is drawn");
    line.split("dy: ")
        .nth(1)
        .and_then(|rest| rest.split("pt").next())
        .and_then(|value| value.trim().parse::<f64>().ok())
        .expect("the label is placed")
}

fn bottom_legend_chart(names: &[&str]) -> Chart {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Bar;
    chart.categories = vec!["Q1".to_string(), "Q2".to_string()];
    chart.has_legend = true;
    chart.legend_position = LegendPosition::Bottom;
    chart.series.truncate(names.len().min(chart.series.len()));
    for (series, name) in chart.series.iter_mut().zip(names) {
        series.name = Some((*name).to_string());
        series.values = vec![4.0, 8.0];
    }
    chart
}

#[test]
fn a_long_legend_name_pushes_the_next_entry_clear() {
    // Every entry advanced by a flat 78pt, so a name wider than that ran under
    // the entry beside it and the two overprinted.
    let short: String = chart_source(bottom_legend_chart(&["A", "B"]));
    let long: String = chart_source(bottom_legend_chart(&[
        "A considerably longer series name",
        "B",
    ]));
    let short_gap: f64 = legend_entry_x(&short, "B") - legend_entry_x(&short, "A");
    let long_gap: f64 =
        legend_entry_x(&long, "B") - legend_entry_x(&long, "A considerably longer series name");
    assert!(
        long_gap > short_gap,
        "a wide name must push its neighbour further along: {long_gap} against {short_gap}"
    );
}

#[test]
fn short_legend_names_keep_the_calibrated_pitch() {
    // The measured width is floored at the old constant, so a legend of short
    // names lays out exactly where it always did.
    let source: String = chart_source(bottom_legend_chart(&["A", "B"]));
    let gap: f64 = legend_entry_x(&source, "B") - legend_entry_x(&source, "A");
    assert!(
        (gap - LEGEND_ENTRY_W).abs() < 1e-9,
        "short names keep the {}pt pitch, got {gap}",
        format_f64(LEGEND_ENTRY_W)
    );
}

#[test]
fn an_excel_legend_entry_keeps_excels_trailing_clearance() {
    // Native Excel for Mac 16.112 exports of `Gift Budget and Tracker1.xlsx`
    // put the second key 96.221pt after the first at 9pt Calibri: 19.2pt of
    // key, 2.025pt before the label, the label's 59.071pt design advance, and
    // 15.930pt of trailing clearance. The generic 6pt `GAP` leaves the second
    // entry almost 10pt too far left. The clearance is the row's, so the row
    // keeps all four of the workbook's names.
    let source: String = chart_source(excel_gift_budget_legend_chart("Calibri", 9.0));
    let pitch: f64 =
        legend_entry_x(&source, "Holiday Budget") - legend_entry_x(&source, "Birthday Budget");

    assert!(
        (pitch - 96.221).abs() <= 0.02,
        "the 9pt Calibri entry pitch is {pitch}pt, Excel's is 96.221pt"
    );
}

#[test]
fn an_excel_legend_centres_the_visible_row() {
    // Across eight faces and every Segoe UI size from 3 through 22pt, native
    // Excel centres the visible row from its first key through its last label;
    // the trailing clearance after that label does not pull the row left. This
    // IR holds one chart-wide family, so restating the legend in Calibri also
    // restates the value labels and moves the content centre 1.022pt left of
    // the legend-only native probe. The real Segoe UI fixture is the exact
    // native-origin gate; this helper isolates the row-centering calculation.
    let mut chart = combo_budget_chart();
    let mut other = chart.series[0].clone();
    other.name = Some("Other Gift Budget".to_string());
    other.values = vec![0.0, 0.0];
    chart.series.insert(2, other);
    chart.series.last_mut().expect("the line series").values = vec![30.0, 180.0];
    chart.text_font_family = Some("Calibri".to_string());
    chart.legend_text_style.size_pt = Some(9.0);
    chart.value_axis_text_style.size_pt = Some(9.0);
    chart.value_axis_number_format = Some("\\$#,##0".to_string());
    let source = framed_chart_source(&chart, EXCEL_GIFT_CHART_FRAME.0, EXCEL_GIFT_CHART_FRAME.1);
    let first_key_x = legend_entry_x(&source, "Birthday Budget");

    assert!(
        (first_key_x - 326.142).abs() <= 0.02,
        "the first key starts at {first_key_x}pt, the calibrated restatement at 326.142pt"
    );
}

#[test]
fn an_excel_legend_gutter_tracks_its_face_and_size() {
    // Fresh one-factor native Excel 16.112 exports at both 9pt and 18pt. Each
    // row was compared with a layout-identical re-zip control; the stated
    // gutter is the entry pitch after subtracting the 19.2pt key, 2.025pt gap,
    // and the source face's design advances. The rows keep the workbook's four
    // names, so each face's clearance is 0.16 of its own mean label advance
    // over the fixed part: no per-face table is needed to reproduce them.
    let names: Vec<String> = [
        "Birthday Budget",
        "Holiday Budget",
        "Other Gift Budget",
        "Amount Spent",
    ]
    .iter()
    .map(|name| (*name).to_string())
    .collect();
    // Each row also carries the source face's own design advance of `Birthday
    // Budget` in its em. A runner without that face installed resolves a
    // substitute whose labels measure differently, and the clearance cannot be
    // asserted against that input; Calibri's advances are calibrated in-tree
    // and reach every runner, which the assertion after the loop holds.
    let measurements = [
        ("Calibri", 13442.0 / 2048.0, 15.930_f64, 25.217_f64),
        ("Arial", 14573.0 / 2048.0, 16.704, 26.788),
        ("Georgia", 14820.0 / 2048.0, 16.847, 27.093),
        ("Times New Roman", 13595.0 / 2048.0, 15.998, 25.379),
        ("Verdana", 16660.0 / 2048.0, 18.058, 29.533),
        ("Century Gothic", 15839.0 / 2048.0, 17.664, 28.718),
        ("Aptos", 14042.0 / 2048.0, 16.366, 26.136),
    ];
    let mut checked: Vec<&str> = Vec::new();
    for (family, source_label_em, at_nine, at_eighteen) in measurements {
        let Some(resolved_label_em) = chart_text_advance_em(family, false, "Birthday Budget")
        else {
            continue;
        };
        if (resolved_label_em - source_label_em).abs() > 1e-6 {
            continue;
        }
        checked.push(family);
        for (size_pt, expected) in [(9.0, at_nine), (18.0, at_eighteen)] {
            let chart = excel_bottom_legend_chart(family, size_pt);
            let actual = excel_legend_trailing_gutter_pt(&chart, &names)
                .expect("a worksheet axis chart has Excel's gutter");
            assert!(
                (actual - expected).abs() <= 0.05,
                "{family} {size_pt}pt leaves {actual}pt, Excel's {expected}pt"
            );
        }
    }
    assert!(
        checked.contains(&"Calibri"),
        "the calibrated Calibri row must be checked on every runner; checked {checked:?}"
    );
}

#[test]
fn non_calibrated_hosts_and_variants_keep_the_generic_legend_gutter() {
    let names: Vec<String> = vec!["Birthday Budget".to_string()];
    for host in [
        crate::ir::ChartHost::Presentation,
        crate::ir::ChartHost::WordProcessing,
        crate::ir::ChartHost::SpreadsheetChartsheet,
    ] {
        let mut chart = excel_bottom_legend_chart("Calibri", 9.0);
        chart.host = host;
        assert_eq!(excel_legend_trailing_gutter_pt(&chart, &names), None);
    }

    let mut line_chart = excel_bottom_legend_chart("Calibri", 9.0);
    line_chart.chart_type = ChartType::Line;
    assert_eq!(excel_legend_trailing_gutter_pt(&line_chart, &names), None);
}

/// A two-entry Segoe UI worksheet legend whose series share one name, the
/// way the #1616 label-width probes rename all four fixture series at once.
fn excel_uniform_legend_chart(size_pt: f64, name: &str) -> Chart {
    let mut chart = excel_bottom_legend_chart("Segoe UI", size_pt);
    for series in &mut chart.series {
        series.name = Some(name.to_string());
    }
    chart
}

/// The four-entry bottom legend of the public #1603 workbook restated in
/// `family` at `size_pt`: three column series and the overlaid line series.
fn excel_gift_budget_legend_chart(family: &str, size_pt: f64) -> Chart {
    let mut chart = excel_bottom_legend_chart(family, size_pt);
    let extra = chart.series[1].clone();
    chart.series.push(extra.clone());
    chart.series.push(extra);
    for (series, name) in chart.series.iter_mut().zip([
        "Birthday Budget",
        "Holiday Budget",
        "Other Gift Budget",
        "Amount Spent",
    ]) {
        series.name = Some(name.to_string());
    }
    chart.series[3].plot_type = Some(ChartType::Line);
    chart
}

#[test]
fn an_excel_legend_entry_advances_by_its_own_label_with_no_width_floor() {
    // Native Excel for Mac 16.112 exports of the #1603 workbook with all four
    // series renamed to one string, twelve widths at 9pt, five at 6pt and four
    // at 18pt, each series guarded by a layout-identical re-zip control
    // (`assets/validation/issue-1616/`). Every key lands key + gap + label +
    // clearance after the previous one; nothing holds a short entry at the
    // 78pt floor that put the 6pt fixture row 12pt left of native (#1616).
    let measurements = [
        (9.0_f64, "It", 34.156_f64),
        (9.0, "Gift", 44.335),
        (9.0, "Spent", 54.427),
        (9.0, "Gift Budget", 80.382),
        (9.0, "Gift Budget Amount", 120.443),
        (6.0, "Gift", 38.843),
        (6.0, "Budget", 49.973),
        (6.0, "Holiday Budget", 75.290),
        (18.0, "Gift", 60.836),
        (18.0, "Budget", 94.215),
    ];
    for (size_pt, name, native_pitch) in measurements {
        let source: String = chart_source(excel_uniform_legend_chart(size_pt, name));
        let marker: String = format!("[{name}]])");
        let entry_x: Vec<f64> = source
            .lines()
            .filter(|line| line.contains(&marker))
            .map(|line| {
                line.split("dx: ")
                    .nth(1)
                    .and_then(|rest| rest.split("pt").next())
                    .and_then(|value| value.trim().parse::<f64>().ok())
                    .expect("the entry is placed")
            })
            .collect();
        assert_eq!(entry_x.len(), 2, "both entries are drawn:\n{source}");
        let pitch: f64 = entry_x[1] - entry_x[0];
        assert!(
            (pitch - native_pitch).abs() <= 0.05,
            "{size_pt}pt {name:?}: entries {pitch}pt apart, Excel's {native_pitch}pt"
        );
    }
}

#[test]
fn an_excel_legend_row_shares_one_clearance_from_its_mean_label() {
    // In the unchanged #1603 workbook the three column labels differ by 9pt of
    // width, yet the native export trails each by the same 16.74pt: the
    // clearance is 0.16 of the row's MEAN label advance, not of each entry's
    // own. Restating only the legend size to 6pt (the #1616 controlled
    // package) keeps the rule and moves every pitch off the old 78pt floor.
    let measurements = [
        (9.0_f64, [101.895_f64, 99.305, 108.495]),
        (6.0, [77.226, 75.500, 81.631]),
    ];
    for (size_pt, native_pitches) in measurements {
        let source: String = chart_source(excel_gift_budget_legend_chart("Segoe UI", size_pt));
        let names = [
            "Birthday Budget",
            "Holiday Budget",
            "Other Gift Budget",
            "Amount Spent",
        ];
        let entry_x: Vec<f64> = names
            .iter()
            .map(|name| legend_entry_x(&source, name))
            .collect();
        for (index, native_pitch) in native_pitches.into_iter().enumerate() {
            let pitch: f64 = entry_x[index + 1] - entry_x[index];
            assert!(
                (pitch - native_pitch).abs() <= 0.05,
                "{size_pt}pt: entry {} is {pitch}pt after entry {index}, Excel's {native_pitch}pt",
                index + 1
            );
        }
    }
}

/// Rebuild the fixture package with one part rewritten, the way the native
/// one-factor probes derive their controlled packages.
fn repackaged_with_part(data: &[u8], part: &str, patch: impl Fn(&str) -> String) -> Vec<u8> {
    let mut archive =
        zip::ZipArchive::new(std::io::Cursor::new(data)).expect("the fixture is a ZIP");
    let mut rebuilt = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let mut patched: bool = false;
    for index in 0..archive.len() {
        let mut entry = archive.by_index(index).expect("the entry is readable");
        let name: String = entry.name().to_string();
        let mut bytes: Vec<u8> = Vec::new();
        std::io::Read::read_to_end(&mut entry, &mut bytes).expect("the entry decompresses");
        rebuilt
            .start_file(&name, zip::write::FileOptions::default())
            .expect("the entry is writable");
        if name == part {
            let xml: String = String::from_utf8(bytes).expect("the part is UTF-8 XML");
            let replacement: String = patch(&xml);
            assert_ne!(replacement, xml, "the patch must change {part}");
            patched = true;
            std::io::Write::write_all(&mut rebuilt, replacement.as_bytes())
                .expect("the patched part is writable");
        } else {
            std::io::Write::write_all(&mut rebuilt, &bytes).expect("the entry is writable");
        }
    }
    assert!(patched, "{part} is in the package");
    rebuilt.finish().expect("the package closes").into_inner()
}

/// The #1616 controlled package: the public workbook with only the legend's
/// `a:defRPr/@sz` changed from 900 to 600. Native Excel for Mac 16.112 keeps
/// its row on the same rule and starts it 38.74 printed points further right
/// than the 9pt row; the filled keys pin both the entry pitch and the visible
/// row's centre after parsing and compiling the real package.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn gift_budget_compiled_six_point_legend_keys_match_native_bounds() {
    let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1603_gift_budget.xlsx");
    let package: Vec<u8> = repackaged_with_part(data, "xl/charts/chart1.xml", |xml| {
        let legend_start: usize = xml.find("<c:legend>").expect("the chart has a legend");
        let legend_end: usize = xml.find("</c:legend>").expect("the legend closes");
        let legend: &str = &xml[legend_start..legend_end];
        assert_eq!(legend.matches("sz=\"900\"").count(), 1);
        format!(
            "{}{}{}",
            &xml[..legend_start],
            legend.replacen("sz=\"900\"", "sz=\"600\"", 1),
            &xml[legend_end..]
        )
    });
    let (doc, _) = crate::parser::Parser::parse(
        &crate::parser::xlsx::XlsxParser,
        &package,
        &crate::config::ConvertOptions::default(),
    )
    .expect("the controlled package parses");
    let output = generate_typst(&doc).unwrap();
    let pages = crate::render::pdf::compiled_page_paint_sequences(&output.source, &output.images)
        .expect("the controlled package compiles");
    assert_eq!(pages.len(), 2);
    let keys: Vec<_> = pages[1]
        .iter()
        .filter(|paint| {
            let (left, top, right, bottom) = paint.bounds;
            paint.rectangle_fill.is_some()
                && (right - left - 15.744).abs() < 0.01
                && top > 350.0
                && bottom < 365.0
        })
        .collect();
    let native = [
        (586.9499, 357.7169, 602.6939, 360.6616),
        (650.2749, 357.7169, 666.0189, 360.6616),
        (712.1853, 357.7169, 727.9293, 360.6616),
    ];
    assert_eq!(keys.len(), native.len());
    for (key, (left, top, right, bottom)) in keys.iter().zip(native) {
        let actual = key.bounds;
        assert!(
            (actual.0 - left).abs() < 0.5
                && (actual.1 - top).abs() < 0.5
                && (actual.2 - right).abs() < 0.5
                && (actual.3 - bottom).abs() < 0.5,
            "native key {:?}, actual {actual:?}",
            (left, top, right, bottom)
        );
        assert!(((actual.3 - actual.1) - (bottom - top)).abs() < 0.01);
    }
}

/// The data-table fallback prints each value through the format its series
/// declares, so a ratio stored as a fraction reads as the percentage the
/// source shows (issue #865).
#[test]
fn test_data_table_prints_a_series_number_format() {
    let chart = Chart {
        // A bubble chart has no plot renderer, so it takes the data-table
        // fallback this rule lives in.
        chart_type: ChartType::Other("bubbleChart".to_string()),
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["Q1".to_string(), "Q2".to_string()],
        series: vec![ChartSeries {
            name: Some("Rate".to_string()),
            values: vec![0.024, 0.689],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: Some("0.0%".to_string()),
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    };
    let source = chart_source(chart);

    assert!(source.contains("2.4%"), "expected 2.4% in: {source}");
    assert!(source.contains("68.9%"), "expected 68.9% in: {source}");
    assert!(
        !source.contains("[0.024]"),
        "the raw fraction must go: {source}"
    );
}

/// A different code is honoured, so the renderer is not special-casing
/// percentages.
#[test]
fn test_data_table_prints_a_declared_thousands_format() {
    let chart = Chart {
        // A bubble chart has no plot renderer, so it takes the data-table
        // fallback this rule lives in.
        chart_type: ChartType::Other("bubbleChart".to_string()),
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["Q1".to_string()],
        series: vec![ChartSeries {
            name: Some("Revenue".to_string()),
            values: vec![1234567.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: Some("#,##0".to_string()),
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    };
    let source = chart_source(chart);

    assert!(
        source.contains("1,234,567"),
        "expected grouping in: {source}"
    );
}

/// A series that states no format keeps the plain rendering, so nothing else
/// in the table moves.
#[test]
fn test_data_table_without_a_number_format_prints_plainly() {
    let chart = Chart {
        // A bubble chart has no plot renderer, so it takes the data-table
        // fallback this rule lives in.
        chart_type: ChartType::Other("bubbleChart".to_string()),
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["Q1".to_string()],
        series: vec![ChartSeries {
            name: Some("Rate".to_string()),
            values: vec![0.024],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    };
    let source = chart_source(chart);

    assert!(
        source.contains("0.024"),
        "expected the plain value in: {source}"
    );
}

/// A currency format emits `$`, which opens math mode in Typst markup. Writing
/// a formatted axis label unescaped produced 48 "unclosed delimiter" errors on
/// a budget workbook in the bulk corpus, so every formatted label is escaped.
#[test]
fn test_a_currency_axis_label_is_escaped() {
    let chart = Chart {
        chart_type: ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["Q1".to_string()],
        series: vec![ChartSeries {
            name: Some("Spend".to_string()),
            values: vec![1200.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: Some("\"$\"#,##0".to_string()),
        secondary_value_axis: None,
        auto_title_deleted: false,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    };
    let source = chart_source(chart);

    assert!(
        source.contains("\\$"),
        "a currency tick label must be escaped: {source}"
    );
    assert!(
        !source.contains("[$"),
        "an unescaped $ opens math mode: {source}"
    );
}

/// A single-series chart that declines the automatic title must not get one
/// from its series name (issue #883).
fn single_series_chart(auto_title_deleted: bool) -> Chart {
    Chart {
        chart_type: ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["Q1".to_string()],
        series: vec![ChartSeries {
            name: Some("Serie 1".to_string()),
            values: vec![1.0],
            fill: None,
            fill_mode: crate::ir::ChartFillMode::Automatic,
            point_fill_modes: Vec::new(),
            point_fills: Vec::new(),
            data_labels: DataLabels::default(),
            number_format: None,
            plot_type: None,
            value_axis: crate::ir::ChartValueAxisRole::Primary,
            marker_symbol: None,
            marker_style: Default::default(),
            line_width_pt: None,
            line_geometry: Default::default(),
        }],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Right,
        has_legend: false,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::Outside,
        value_axis_major_tick_mark: AxisTickMark::Outside,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

#[test]
fn test_auto_title_deleted_suppresses_the_series_name_title() {
    let source = chart_source(single_series_chart(true));
    assert!(
        !source.contains("Serie 1"),
        "a declined automatic title must not be drawn: {source}"
    );
}

/// Triangulation: the fallback itself stays, so a chart that does not decline
/// it still gets its automatic title.
#[test]
fn test_a_chart_that_keeps_its_automatic_title_still_gets_one() {
    let source = chart_source(single_series_chart(false));
    assert!(
        source.contains("Serie 1"),
        "the automatic title must survive: {source}"
    );
}

// ----- Automatic chart title (issue #1146) -----

/// A `<c:title>` naming no text is Office's automatic title: the part carries
/// the formatting and the application supplies the string.
///
/// `tests/fixtures/xlsx/any_sheets.xlsx` writes one over two unnamed series.
/// An Excel for Mac 16.100 export of its `Chart` chartsheet, forced to Letter
/// landscape, prints the placeholder centred at the top of the chart box and
/// starts the plot below it: the chart box runs y 58..551.18 and the topmost
/// gridline sits at y 102.20, where without a title it would sit within about
/// 10pt of the box top.
fn automatic_title_chart(series_names: &[Option<&str>], auto_title_deleted: bool) -> Chart {
    Chart {
        chart_type: ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["Q1".to_string(), "Q2".to_string(), "Q3".to_string()],
        series: series_names
            .iter()
            .map(|name| ChartSeries {
                name: name.map(str::to_string),
                values: vec![1.0, 3.0, 5.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            })
            .collect(),
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Bottom,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::None,
        value_axis_major_tick_mark: AxisTickMark::None,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::default(),
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted,
        has_automatic_title: true,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

#[test]
fn an_automatic_title_over_unnamed_series_prints_the_placeholder() {
    let source = chart_source(automatic_title_chart(&[None, None], false));

    assert!(
        source.contains("Chart Title"),
        "an automatic title must be drawn: {source}"
    );
}

/// The placeholder is only for charts the series cannot name: a lone named
/// series still lends its name, exactly as it did before (issue #883).
#[test]
fn a_lone_named_series_still_outranks_the_placeholder() {
    let source = chart_source(automatic_title_chart(&[Some("Serie 1")], false));

    assert!(
        source.contains("Serie 1"),
        "the series name is the automatic title where there is one: {source}"
    );
    assert!(
        !source.contains("Chart Title"),
        "the placeholder must not double the series name: {source}"
    );
}

/// Two named series cannot name the chart between them, so Office falls back
/// to the placeholder there too.
#[test]
fn two_named_series_still_get_the_placeholder() {
    let source = chart_source(automatic_title_chart(&[Some("Left"), Some("Right")], false));

    assert!(
        source.contains("Chart Title"),
        "a multi-series chart takes the placeholder: {source}"
    );
}

/// `<c:autoTitleDeleted val="1"/>` declines the automatic title whatever the
/// application would have supplied for it.
#[test]
fn a_declined_automatic_title_prints_no_placeholder() {
    let source = chart_source(automatic_title_chart(&[None, None], true));

    assert!(
        !source.contains("Chart Title"),
        "a declined automatic title must not be drawn: {source}"
    );
}

/// A chart declaring no `<c:title>` gets nothing — the placeholder follows the
/// element, not the absence of a title.
#[test]
fn a_chart_declaring_no_title_element_prints_no_placeholder() {
    let mut chart = automatic_title_chart(&[None, None], false);
    chart.has_automatic_title = false;

    let source = chart_source(chart);

    assert!(
        !source.contains("Chart Title"),
        "no title element means no title: {source}"
    );
}

/// The placeholder takes its band out of the plot, so the plot starts below it
/// rather than running the whole frame the way it does untitled.
#[test]
fn an_automatic_title_shortens_the_plot_by_its_band() {
    let titled = automatic_title_chart(&[None, None], false);
    let mut untitled = automatic_title_chart(&[None, None], false);
    untitled.has_automatic_title = false;

    let titled_source: String = framed_chart_source(&titled, 400.0, 300.0);
    let untitled_source: String = framed_chart_source(&untitled, 400.0, 300.0);
    let band: f64 = chart_area_title_h(&titled);

    assert_eq!(
        block_height(&titled_source, "#block(width: 400pt, height: "),
        Some(band),
        "the title takes a {band}pt band: {titled_source}"
    );
    assert_eq!(
        block_height(&untitled_source, "#block(width: 400pt, height: "),
        None,
        "an untitled chart emits no band at all: {untitled_source}"
    );

    // A titled chart now opens the full chart-area box first and the shortened
    // content box inside it. The last matching box is therefore the one whose
    // height must still give up the band (issue #1216).
    let box_height = |source: &str| last_block_height(source, "#box(width: 400pt, height: ");
    assert_eq!(
        box_height(&untitled_source)
            .zip(box_height(&titled_source))
            .map(|(untitled, titled)| untitled - titled),
        Some(band),
        "and the plot box gives up exactly that much"
    );

    let (_, _, _, titled_plot) = plot_rect(&emitted_lines(&titled_source));
    let (_, _, _, untitled_plot) = plot_rect(&emitted_lines(&untitled_source));
    assert!(
        titled_plot < untitled_plot,
        "so the plotting rectangle itself is shorter: {titled_plot}pt against {untitled_plot}pt"
    );
}

/// The height of the first block or box the source opens with `prefix`.
fn block_height(source: &str, prefix: &str) -> Option<f64> {
    source
        .lines()
        .find_map(|line| line.strip_prefix(prefix))
        .and_then(leading_pt)
}

/// The height of the last block or box the source opens with `prefix`.
fn last_block_height(source: &str, prefix: &str) -> Option<f64> {
    source
        .lines()
        .filter_map(|line| line.strip_prefix(prefix))
        .filter_map(leading_pt)
        .next_back()
}

/// The axes and gridlines draw with the line the part declares, and fall back
/// to the automatic stroke only where it declares none (issue #900).
#[test]
fn declared_axis_and_gridline_lines_reach_the_generated_source() {
    let white = crate::ir::ChartLine::Explicit {
        alpha: None,
        width_pt: Some(1.0),
        color: Some(Color::new(0xFF, 0xFF, 0xFF)),
    };
    let mut chart = stacked_support_chart(ChartGrouping::Clustered);
    chart.category_axis_line = white;
    chart.value_axis_line = white;
    chart.major_gridline_line = white;

    let source = chart_source(chart);

    assert!(
        source.contains("stroke: 1pt + rgb(255, 255, 255)"),
        "the declared white line must reach the source:\n{source}"
    );
    // The chart-area outline is a separate declaration and this chart states
    // none, so it keeps the automatic stroke. Only the `line(...)` draws — the
    // axes, ticks and gridlines — are this issue's.
    let automatic_lines = source
        .lines()
        .filter(|line| line.contains("line(end:") && line.contains("rgb(134, 134, 134)"))
        .count();
    assert_eq!(
        automatic_lines, 0,
        "no axis, tick or gridline should still take the automatic grey:\n{source}"
    );
}

/// A suppressed line draws nothing at all — not the automatic one.
#[test]
fn a_suppressed_axis_line_is_not_drawn() {
    let mut chart = stacked_support_chart(ChartGrouping::Clustered);
    chart.value_axis_line = crate::ir::ChartLine::Suppressed;
    chart.category_axis_line = crate::ir::ChartLine::Suppressed;
    chart.major_gridline_line = crate::ir::ChartLine::Suppressed;

    let source = chart_source(chart);

    let drawn_lines = source
        .lines()
        .filter(|line| line.contains("line(end:"))
        .count();
    assert_eq!(
        drawn_lines, 0,
        "a suppressed axis and gridline draw nothing:\n{source}"
    );
}

/// A chart that declares nothing keeps the automatic stroke, so the fallback
/// is not lost along the way.
#[test]
fn an_undeclared_axis_still_draws_the_automatic_line() {
    let source = chart_source(stacked_support_chart(ChartGrouping::Clustered));

    assert!(
        source.contains("0.75pt + rgb(134, 134, 134)"),
        "the automatic stroke must survive:\n{source}"
    );
}

/// A clustered column's labels sit beyond the bar's end, and a stacked one's
/// centre on the segment (issue #901).
///
/// The label's `dy` is what moves: `outEnd` puts its box just above the bar
/// top, `ctr` halfway down the bar.
#[test]
fn a_clustered_bar_label_sits_outside_its_end_and_a_stacked_one_centres() {
    fn label_dy(grouping: ChartGrouping, position: crate::ir::DataLabelPosition) -> f64 {
        let mut chart = stacked_support_chart(grouping);
        for series in &mut chart.series {
            series.data_labels = DataLabels {
                show_value: true,
                position,
                position_stated: true,
                ..DataLabels::default()
            };
        }
        let source = chart_source(chart);
        // The first label draw after the first bar rect.
        source
            .lines()
            .filter(|line| line.contains("align(center + horizon)"))
            .find_map(|line| {
                let dy = line.split("dy: ").nth(1)?;
                dy.split("pt").next()?.parse::<f64>().ok()
            })
            .expect("a data label is drawn")
    }

    let centred = label_dy(ChartGrouping::Stacked, crate::ir::DataLabelPosition::Center);
    let outside = label_dy(
        ChartGrouping::Clustered,
        crate::ir::DataLabelPosition::OutsideEnd,
    );

    assert!(
        outside < centred,
        "an outEnd label must sit above a centred one: outEnd dy {outside}, ctr dy {centred}"
    );
}

/// An `outEnd` label clears the bar's end rather than sitting flush against it
/// (issue #907).
///
/// The reference leaves about 2.8pt whatever the label's size — 8pt labels
/// clear by a mean 2.66pt, 11.97pt by 2.99pt, 18pt by 2.73pt — so the
/// placement carries a constant offset, asserted here against the bar's own
/// `dy` in the same generated source.
#[test]
fn an_outside_end_label_clears_the_bar_by_a_constant() {
    let mut chart = stacked_support_chart(ChartGrouping::Clustered);
    for series in &mut chart.series {
        series.data_labels = DataLabels {
            show_value: true,
            position: crate::ir::DataLabelPosition::OutsideEnd,
            position_stated: true,
            ..DataLabels::default()
        };
    }
    let source = chart_source(chart);

    fn first_dy(source: &str, needle: &str) -> f64 {
        source
            .lines()
            .filter(|line| line.contains(needle))
            .find_map(|line| line.split("dy: ").nth(1)?.split("pt").next()?.parse().ok())
            .unwrap_or_else(|| panic!("no {needle} draw in:\n{source}"))
    }

    // The first bar rect and the first label box, in draw order.
    let bar_top: f64 = first_dy(&source, "rect(width:");
    let label_dy: f64 = first_dy(&source, "align(center + horizon)");

    let clearance: f64 = bar_top - label_dy;
    assert!(
        clearance > 10.0,
        "an outEnd label must clear the bar top by more than its own line box, \
         got {clearance} (bar {bar_top}, label {label_dy})"
    );
    assert!(
        (clearance - 12.4).abs() < 0.01,
        "expected the 10pt line box plus the 2.4pt gap, got {clearance}"
    );
}

/// A declared `<c:majorUnit>` sets the tick interval (issue #882).
///
/// The deck in #841 declares 0.2 on a 0.689 maximum and the reference ticks
/// 0/20/40/60/80%; we ticked every 10%, twice as often as the file asks.
#[test]
fn a_stated_major_unit_sets_the_tick_interval() {
    fn tick_labels(unit: Option<f64>) -> Vec<String> {
        let mut chart = stacked_support_chart(ChartGrouping::Clustered);
        chart.value_axis_major_unit = unit;
        let source = chart_source(chart);
        source
            .lines()
            .filter(|line| line.contains("align(right + horizon)"))
            .filter_map(|line| {
                let start = line.rfind('[')?;
                Some(
                    line[start + 1..]
                        .trim_end_matches(&[']', ')'][..])
                        .to_string(),
                )
            })
            .collect()
    }

    let automatic = tick_labels(None);
    let stated = tick_labels(Some(4.0));

    assert!(
        stated.len() < automatic.len(),
        "a 4-unit interval must give fewer ticks than the automatic one: \
         automatic {automatic:?}, stated {stated:?}"
    );
    assert!(
        stated.len() >= 2,
        "the axis still needs its ticks: {stated:?}"
    );
}

// ----- Crowded category labels slant (issue #884) -----

/// A column chart whose category labels are far longer than their bands, as
/// the deck in #841 has.
fn crowded_column_chart() -> Chart {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Column;
    chart.series.truncate(1);
    chart.categories = vec![
        "Fortjenestemargin".to_string(),
        "Bruttofortjeneste".to_string(),
        "Konverteringsfrekvens for kundeemne".to_string(),
        "Frekvens for kundebevaring".to_string(),
    ];
    chart.series[0].values = vec![33.9, 68.9, 2.4, 9.3];
    chart.title = None;
    chart
}

#[test]
fn crowded_category_labels_slant_by_forty_five_degrees() {
    let source: String = chart_source(crowded_column_chart());
    assert!(
        source.contains("rotate(-45deg, origin: top + right"),
        "labels longer than their band must slant, got:\n{source}"
    );
}

#[test]
fn a_powerpoint_rotated_category_label_takes_the_native_baseline_seat() {
    let mut chart = crowded_column_chart();
    chart.host = crate::ir::ChartHost::Presentation;
    chart.text_style.size_pt = Some(11.97);
    chart.category_axis_text_style.size_pt = Some(11.97);
    chart.text_font_family = Some("Avenir Next LT Pro".to_string());

    let axis_y = 200.0;
    let expected = axis_y + 2.0 + 0.74 * 11.97;
    let actual = chart_category_rotated_label_y(&chart, axis_y);
    assert!((actual - expected).abs() < 0.001, "{actual}");
}

#[test]
fn a_non_powerpoint_rotated_category_label_keeps_its_existing_seat() {
    let mut chart = crowded_column_chart();
    chart.host = crate::ir::ChartHost::Spreadsheet;
    chart.text_style.size_pt = Some(11.97);
    chart.category_axis_text_style.size_pt = Some(11.97);
    chart.text_font_family = Some("Avenir Next LT Pro".to_string());

    assert_eq!(chart_category_rotated_label_y(&chart, 200.0), 202.0);
}

/// Issue #1022: with the vertical seat corrected (#1014), all four rotated
/// label runs on the #841 deck's slide 14 still land 3.32–3.35pt right of the
/// fresh native export at 11.97pt while their baselines agree — a uniform
/// 0.279em inset between the band centre and the rotated trailing end.
#[test]
fn a_powerpoint_rotated_category_label_insets_its_trailing_anchor() {
    let mut chart = crowded_column_chart();
    chart.host = crate::ir::ChartHost::Presentation;
    chart.text_style.size_pt = Some(11.97);
    chart.category_axis_text_style.size_pt = Some(11.97);
    chart.text_font_family = Some("Avenir Next LT Pro".to_string());

    let centre = 300.0;
    let label_box_w = 80.0;
    let expected = centre - label_box_w - 0.279 * 11.97;
    let actual = chart_category_rotated_label_x(&chart, centre, label_box_w);
    assert!((actual - expected).abs() < 0.001, "{actual}");
}

/// Issue #1025: with axis, plot, and tick seats calibrated, the four column
/// data labels on the #841 deck still sit 1.32-1.43pt below the native
/// baselines with x agreeing to 0.01pt — a uniform 0.114em seat at the
/// label size.
#[test]
fn a_powerpoint_column_data_label_takes_the_native_seat() {
    let mut chart = crowded_column_chart();
    chart.host = crate::ir::ChartHost::Presentation;
    chart.text_style.size_pt = Some(11.97);

    let labels = crate::ir::DataLabels::default();
    let expected = 0.114 * 11.97;
    let actual = pptx_column_data_label_seat_pt(&chart, &labels);
    assert!((actual - expected).abs() < 0.001, "{actual}");
}

#[test]
fn a_non_powerpoint_column_data_label_keeps_its_seat() {
    let mut chart = crowded_column_chart();
    chart.host = crate::ir::ChartHost::Spreadsheet;
    chart.text_style.size_pt = Some(11.97);

    assert_eq!(
        pptx_column_data_label_seat_pt(&chart, &crate::ir::DataLabels::default()),
        0.0
    );
}

#[test]
fn a_non_powerpoint_rotated_category_label_keeps_its_trailing_anchor() {
    let mut chart = crowded_column_chart();
    chart.host = crate::ir::ChartHost::Spreadsheet;
    chart.text_style.size_pt = Some(11.97);
    chart.category_axis_text_style.size_pt = Some(11.97);
    chart.text_font_family = Some("Avenir Next LT Pro".to_string());

    assert_eq!(
        chart_category_rotated_label_x(&chart, 300.0, 80.0),
        300.0 - 80.0
    );
}

#[test]
fn crowded_category_labels_keep_declared_character_spacing_without_moving_the_plot() {
    // Slide 14 of the #841 deck declares `spc="100"` on both axes. Dropping
    // that 1pt gap shortened the longest category by 34pt and pulled every
    // right-aligned diagonal label up and right in the rendered comparison.
    let mut chart = crowded_column_chart();
    chart.text_style.size_pt = Some(11.97);
    chart.text_font_family = Some("Avenir Next LT Pro".to_string());
    chart.category_axis_text_style.letter_spacing_hundredths = Some(100);

    let tracked_plot = axis_plot_rect(&chart, (401.95, 344.25), false);
    let source = chart_source(chart);
    assert!(
        source.contains(", tracking: 1pt, ligatures: false, kerning: false)[Fortjenestemargin]"),
        "axis text must emit the declared spacing: {source}"
    );
    let mut plain = crowded_column_chart();
    plain.text_style.size_pt = Some(11.97);
    plain.text_font_family = Some("Avenir Next LT Pro".to_string());
    let plain_plot = axis_plot_rect(&plain, (401.95, 344.25), false);
    assert_eq!(
        tracked_plot, plain_plot,
        "tracking changes glyph advance, not the native-calibrated chart plot"
    );
}

#[test]
fn a_crowded_axis_with_ellipsis_overflow_shortens_only_the_label_that_exceeds_its_box() {
    let mut chart = crowded_column_chart();
    chart.text_style.size_pt = Some(11.97);
    chart.text_font_family = Some("Avenir Next LT Pro".to_string());
    chart.category_axis_text_style.letter_spacing_hundredths = Some(100);
    chart.category_axis_text_style.ellipsis_overflow = true;

    let source = chart_source(chart);
    // Issue #1076: only the retained text right-aligns at the trailing-end
    // anchor. The swallowed inter-word space still paints — a 0.25em space at
    // 11.97pt, one 1pt tracking step past the anchor — but is pulled back out
    // of the aligned width, and the ellipsis takes the place of the first
    // character it replaced, one tracking step past the anchor.
    assert!(
        source.contains(
            "[Konverteringsfrekvens for#\" \";#h(-3.9925pt);\
             #box(width: 0pt)[#align(left)[#move(dx: 1pt)[…]]]]"
        ),
        "only the retained text aligns; space and ellipsis hang: {source}"
    );
    assert!(
        !source.contains("[Konverteringsfrekvens for kundeemne]"),
        "the hidden suffix must not be painted: {source}"
    );
    assert!(
        source.contains("[Frekvens for kundebevaring]"),
        "a label that still fits must remain complete: {source}"
    );
}

#[test]
fn an_ellipsized_rotated_label_hangs_by_its_own_space_and_tracking() {
    // Triangulation for #1076: doubling the declared spacing must move both the
    // pull-back and the ellipsis with it, or the offsets are baked-in constants
    // rather than the label's own metrics.
    let mut chart = crowded_column_chart();
    chart.text_style.size_pt = Some(11.97);
    chart.text_font_family = Some("Avenir Next LT Pro".to_string());
    chart.category_axis_text_style.letter_spacing_hundredths = Some(200);
    chart.category_axis_text_style.ellipsis_overflow = true;

    let source = chart_source(chart);
    assert!(
        source.contains("#h(-4.9925pt);#box(width: 0pt)[#align(left)[#move(dx: 2pt)[…]]]"),
        "the hang follows the declared 2pt tracking and the 2.9925pt space: {source}"
    );
}

#[test]
fn a_rotated_label_truncated_mid_word_hangs_only_its_ellipsis() {
    // Triangulation for #1076: with no inter-word space to swallow there is
    // nothing to pull back, and the ellipsis alone overflows the anchor.
    let mut chart = crowded_column_chart();
    chart.categories[2] = "Konverteringsfrekvensforkundeemne".to_string();
    chart.text_style.size_pt = Some(11.97);
    chart.text_font_family = Some("Avenir Next LT Pro".to_string());
    chart.category_axis_text_style.letter_spacing_hundredths = Some(100);
    chart.category_axis_text_style.ellipsis_overflow = true;

    let source = chart_source(chart);
    assert!(
        source.contains("#box(width: 0pt)[#align(left)[#move(dx: 1pt)[…]]]"),
        "the ellipsis still hangs one tracking step past the anchor: {source}"
    );
    assert!(
        !source.contains("#h(-"),
        "nothing was swallowed, so nothing is pulled back: {source}"
    );
}

#[test]
fn category_labels_that_fit_their_band_stay_flat() {
    // Triangulation: the same chart type with short labels must not rotate,
    // or the rule is "always rotate" rather than "rotate when crowded".
    let mut chart = crowded_column_chart();
    chart.categories = vec![
        "Q1".to_string(),
        "Q2".to_string(),
        "Q3".to_string(),
        "Q4".to_string(),
    ];
    let source: String = chart_source(chart);
    assert!(
        !source.contains("rotate(-45deg"),
        "short labels must stay flat, got:\n{source}"
    );
}

// ----- A chart's declared text colour (issue #916) -----

#[test]
fn axis_labels_take_the_declared_chart_text_colour() {
    // The deck in #841 sets its chart text white against a dark chart area;
    // the tick labels printed black because no colour was ever parsed.
    let mut chart = sized_bar_chart(11.0);
    chart.text_style.color = Some(crate::ir::Color::new(255, 255, 255));
    let source: String = chart_source(chart);
    assert!(
        source.contains("#text(size: 11pt, fill: rgb(255, 255, 255))[Q1]"),
        "a category label must take the declared colour, got:\n{source}"
    );
    assert!(
        source.contains("fill: rgb(255, 255, 255))[0]"),
        "a value tick label must take it too, got:\n{source}"
    );
}

#[test]
fn an_axis_colour_overrides_the_chart_space_colour_for_that_axis_only() {
    let mut chart = sized_bar_chart(11.0);
    chart.text_style.color = Some(crate::ir::Color::new(255, 255, 255));
    chart.category_axis_text_style.color = Some(crate::ir::Color::new(255, 0, 0));
    let source: String = chart_source(chart);
    assert!(
        source.contains("#text(size: 11pt, fill: rgb(255, 0, 0))[Q1]"),
        "the category axis' own colour must win, got:\n{source}"
    );
    assert!(
        source.contains("fill: rgb(255, 255, 255))[0]"),
        "the value axis must keep the chart-space colour, got:\n{source}"
    );
}

#[test]
fn a_chart_declaring_no_colour_keeps_the_colours_it_had() {
    // Triangulation: the fix must not force a default onto every chart. An
    // axis label stays uncoloured and a data label stays the hardcoded white.
    let mut chart = sized_bar_chart(11.0);
    // The white only appears where a data label is drawn at all.
    chart.series[0].data_labels.show_value = true;
    let source: String = chart_source(chart);
    // Bars carry their series colour as a fill and a legend swatch sits on the
    // same line as its label, so the question is only about what is inside a
    // `#text(...)` argument list.
    let coloured_runs: Vec<&str> = source
        .match_indices("#text(")
        .filter_map(|(at, _)| {
            let args = &source[at + "#text(".len()..];
            let args = &args[..args.find(')')?];
            // `fill: white` is the data label's long-standing default and is
            // asserted separately below; only a resolved colour would be new.
            args.contains("fill: rgb(").then_some(args)
        })
        .collect();
    assert!(
        coloured_runs.is_empty(),
        "nothing declared must leave text uncoloured, got: {coloured_runs:?}"
    );
    assert!(
        source.contains("fill: white)"),
        "the data label keeps its white, got:\n{source}"
    );
}

/// A data label was written at a literal 8pt, so a chart declaring anything
/// else drew its labels at the wrong size — smaller than its own axis on the
/// deck of #841, which asks for 11.97pt everywhere (issue #970).
#[test]
fn a_data_label_is_set_at_the_size_its_dlbls_declare() {
    let mut chart = labelled_chart(DataLabels {
        show_value: true,
        text_style: crate::ir::ChartTextStyle {
            size_pt: Some(11.97),
            ..crate::ir::ChartTextStyle::default()
        },
        ..DataLabels::default()
    });
    chart.series[0].data_labels.text_style.size_pt = Some(11.97);
    let source = chart_source(chart);

    assert!(
        source.contains("#text(size: 11.97pt, weight: \"bold\""),
        "{source}"
    );
    assert!(
        !source.contains("#text(size: 8pt, weight: \"bold\""),
        "{source}"
    );
    // The label box has to grow with the text or the larger glyphs centre on
    // a box sized for 8pt: 11.97 x 1.25 is 14.9625pt.
    assert!(source.contains("height: 14.9625pt"), "{source}");
}

/// A `<c:dLbls>` stating no size takes the chart space's, and only a chart
/// stating nothing anywhere keeps the unmeasured 8pt the labels were pinned
/// at — reading a declared size must not resize charts that declare none.
#[test]
fn a_data_label_declaring_no_size_falls_back_to_the_chart_space() {
    let mut chart = labelled_chart(DataLabels {
        show_value: true,
        ..DataLabels::default()
    });
    chart.text_style.size_pt = Some(18.0);
    assert!(chart_source(chart).contains("#text(size: 18pt, weight: \"bold\""),);

    let neither = chart_source(labelled_chart(DataLabels {
        show_value: true,
        ..DataLabels::default()
    }));
    assert!(
        neither.contains("#text(size: 8pt, weight: \"bold\""),
        "{neither}"
    );
}

// ----- Combo plot areas: bars with a line over them (issue #1067) -----

/// `Gift Budget and Tracker1.xlsx`'s chart, reduced to two categories: three
/// budget series stacked into one column per month, with the amount actually
/// spent drawn as a line over them.
///
/// Stack totals are 0 and 150; the line runs 25 → 75, so it stays inside the
/// stack rather than above it. Nothing about the numbers can be read off the
/// column heights alone.
fn combo_budget_chart() -> Chart {
    let budget = |name: &str, values: Vec<f64>| ChartSeries {
        name: Some(name.to_string()),
        values,
        fill: None,
        fill_mode: crate::ir::ChartFillMode::Automatic,
        point_fill_modes: Vec::new(),
        point_fills: Vec::new(),
        data_labels: DataLabels::default(),
        number_format: None,
        plot_type: None,
        value_axis: crate::ir::ChartValueAxisRole::Primary,
        marker_symbol: None,
        marker_style: Default::default(),
        line_width_pt: None,
        line_geometry: Default::default(),
    };
    Chart {
        chart_type: ChartType::Column,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["May".to_string(), "Jun".to_string()],
        series: vec![
            budget("Birthday Budget", vec![0.0, 50.0]),
            budget("Holiday Budget", vec![0.0, 100.0]),
            ChartSeries {
                name: Some("Amount Spent".to_string()),
                values: vec![25.0, 75.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: Some(ChartType::Line),
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: None,
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
        ],
        grouping: ChartGrouping::Stacked,
        legend_position: LegendPosition::Bottom,
        has_legend: true,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::None,
        value_axis_major_tick_mark: AxisTickMark::None,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: false,
        bar_band_layout: BarBandLayout {
            gap_width_percent: 150.0,
            overlap_percent: 100.0,
        },
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::Spreadsheet,
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: true,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

/// The points of the first emitted `path(...)`, in placement coordinates.
fn emitted_path_points(source: &str) -> Vec<(f64, f64)> {
    let line: &str = source
        .lines()
        .find(|line| line.contains("curve(stroke:"))
        .unwrap_or_default();
    line.match_indices("(")
        .filter_map(|(index, _)| {
            let body: &str = line[index + 1..].split_once(')')?.0;
            let (x, y) = body.split_once(", ")?;
            Some((
                x.strip_suffix("pt")?.parse().ok()?,
                y.strip_suffix("pt")?.parse().ok()?,
            ))
        })
        .collect()
}

#[test]
fn a_line_series_over_columns_draws_a_line_not_a_column() {
    let source = chart_source(combo_budget_chart());

    // Two bar series over two categories. The line series contributes no bar:
    // every one of its points drew as a polyline segment instead. Its markers
    // are triangles at this series index, so they are no rectangle either.
    assert_eq!(
        emitted_rects(&source).len(),
        4,
        "only the two bar-family series may draw columns; got:\n{source}"
    );
    assert_eq!(
        emitted_path_points(&source).len(),
        2,
        "the line series draws a polyline through its two points; got:\n{source}"
    );
}

#[test]
fn a_line_series_reads_against_the_same_axis_as_the_columns() {
    // The line is not part of the stack: its 75 in June is half of that
    // month's 150 stack, so its point sits midway up the column.
    let source = chart_source(combo_budget_chart());
    let bars = emitted_rects(&source);
    let june_x: f64 = bars.iter().map(|bar| bar.dx).fold(f64::MIN, f64::max);
    let june: Vec<PlacedRect> = bars
        .into_iter()
        .filter(|bar| (bar.dx - june_x).abs() < 0.01)
        .collect();
    let stack_top: f64 = june.iter().map(|bar| bar.dy).fold(f64::MAX, f64::min);
    let baseline: f64 = june
        .iter()
        .map(|bar| bar.dy + bar.height)
        .fold(f64::MIN, f64::max);

    let points = emitted_path_points(&source);
    let (june_point_x, june_point_y) = *points.last().expect("a point per category");

    assert!(
        (june_point_y - (stack_top + baseline) / 2.0).abs() < 0.5,
        "75 of a 150 stack sits midway between the baseline {baseline} and the \
         stack top {stack_top}, got {june_point_y}"
    );
    // And on the same category band as June's column, whose bar is centred in
    // it by `<c:overlap val="100"/>`.
    let band_centre: f64 = june_x + june[0].width / 2.0;
    assert!(
        (june_point_x - band_centre).abs() < 0.5,
        "the point sits at its category's centre {band_centre}, got {june_point_x}"
    );
}

#[test]
fn a_line_series_stays_out_of_the_stack_the_axis_is_scaled_to() {
    // Summing the line into each category's stack would carry the axis to 225
    // and shrink every column by a third.
    let axis_max: f64 = emitted_axis_ticks(&chart_source(combo_budget_chart()))
        .iter()
        .copied()
        .fold(0.0, f64::max);

    assert!(
        (150.0..200.0).contains(&axis_max),
        "the axis covers the 150 stack the columns reach, got {axis_max}"
    );
}

#[test]
fn a_combo_legend_draws_each_series_the_way_its_family_plots_it() {
    let source = chart_source(combo_budget_chart());
    let entry = |name: &str| -> String {
        source
            .lines()
            .find(|line| line.contains(&format!("[{name}]")) && line.contains("#place"))
            .unwrap_or_else(|| panic!("no legend entry for {name} in:\n{source}"))
            .to_string()
    };

    // Excel draws a filled swatch for a bar series and a line-and-marker key
    // for a line one.
    assert!(!entry("Birthday Budget").contains("line(end:"));
    assert!(entry("Amount Spent").contains("line(end:"));
}

#[test]
fn an_overlaid_line_series_draws_the_symbol_it_declares() {
    // The audited workbook's line series names `<c:symbol val="circle"/>`, and
    // the native Excel export draws a filled circle on each of its points and
    // on its legend key. Its index put it on the shape cycle's cross instead
    // (issue #1107).
    let mut chart = combo_budget_chart();
    let line_series: &mut ChartSeries = chart.series.last_mut().expect("the line series");
    line_series.marker_symbol = Some(MarkerSymbol::Circle);
    let source = chart_source(chart);

    // A point per category, plus the legend key.
    let (circles, polygons, _) = marker_shape_counts(&source);
    assert_eq!(
        circles, 3,
        "the line series draws a circle on both points and on its legend key; \
         got:\n{source}"
    );
    assert_eq!(
        polygons, 0,
        "no marker may fall back to the cycle's shape for this index; got:\n{source}"
    );
}

// ----- Combo plot areas: a scatter marker over a line (issue #1123) -----

/// The cash-flow chart of `Monthly college budget1.xlsx`, reduced to three
/// months: `<c:lineChart>` draws the month-by-month cash flow and a
/// `<c:scatterChart>` beside it puts one marker on the selected month.
///
/// The scatter series carries a single point because the other twelve are
/// `#N/A` in the workbook's cache — that is how the template hides the marker
/// on every month but the selected one.
fn combo_line_and_scatter_chart() -> Chart {
    Chart {
        chart_type: ChartType::Line,
        hole_size_percent: None,
        first_slice_angle_deg: None,
        title: None,
        categories: vec!["jan".to_string(), "feb".to_string(), "mar".to_string()],
        series: vec![
            ChartSeries {
                name: Some("Cash Flow".to_string()),
                values: vec![169.0, 69.0, 192.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: None,
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: Some(MarkerSymbol::Off),
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
            ChartSeries {
                name: Some("Positive Selected Period".to_string()),
                values: vec![169.0],
                fill: None,
                fill_mode: crate::ir::ChartFillMode::Automatic,
                point_fill_modes: Vec::new(),
                point_fills: Vec::new(),
                data_labels: DataLabels::default(),
                number_format: None,
                plot_type: Some(ChartType::Scatter),
                value_axis: crate::ir::ChartValueAxisRole::Primary,
                marker_symbol: Some(MarkerSymbol::Circle),
                marker_style: Default::default(),
                line_width_pt: None,
                line_geometry: Default::default(),
            },
        ],
        grouping: ChartGrouping::Clustered,
        legend_position: LegendPosition::Bottom,
        has_legend: false,
        category_axis_title: None,
        value_axis_title: None,
        category_axis_major_tick_mark: AxisTickMark::None,
        value_axis_major_tick_mark: AxisTickMark::None,
        category_axis_deleted: false,
        category_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_line: crate::ir::ChartLine::Automatic,
        value_axis_major_unit: None,
        value_axis_min: None,
        value_axis_max: None,
        major_gridline_line: crate::ir::ChartLine::Automatic,
        value_axis_deleted: true,
        bar_band_layout: BarBandLayout::default(),
        theme_accent_colors: Vec::new(),
        chart_area_fill: crate::ir::ChartAreaFill::Unspecified,
        chart_area_outline: ChartAreaOutline::Default,
        host: crate::ir::ChartHost::Spreadsheet,
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
        title_text_style: crate::ir::ChartTextStyle::default(),
        legend_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_text_font_family: None,
        value_axis_text_font_family: None,
        category_axis_text_style: crate::ir::ChartTextStyle::default(),
        value_axis_text_style: crate::ir::ChartTextStyle::default(),
        category_axis_number_format: None,
        value_axis_number_format: None,
        secondary_value_axis: None,
        auto_title_deleted: true,
        has_automatic_title: false,
        title_layout: None,
        plot_area_layout: None,
        user_shapes: Vec::new(),
    }
}

#[test]
fn a_scatter_series_over_a_line_draws_a_plot_not_a_data_table() {
    let source = chart_source(combo_line_and_scatter_chart());

    // The data-table fallback names the chart's family in place of a plot; a
    // drawn chart never prints its own type. Reaching it here cost the
    // workbook a whole extra page (issue #1123).
    assert!(
        !source.contains("Scatter Chart"),
        "a combo with a plottable family must not fall back to a data table; got:\n{source}"
    );
    let points = emitted_path_points(&source);
    assert_eq!(
        points.len(),
        3,
        "the line family draws a polyline through its three months; got:\n{source}"
    );
    // One point per category band, left to right, and February's 69 is the
    // lowest of the three — the plot is scaled to the values, not a table.
    assert!(points[0].0 < points[1].0 && points[1].0 < points[2].0);
    assert!(points[1].1 > points[0].1 && points[0].1 > points[2].1);
}

#[test]
fn a_one_point_scatter_series_draws_its_marker_and_no_polyline() {
    let source = chart_source(combo_line_and_scatter_chart());

    // Its single point is a marker on the selected month, not a segment: a
    // polyline needs two points and the workbook caches `#N/A` for the rest.
    assert_eq!(
        source.matches("curve(stroke:").count(),
        1,
        "only the line family's series draws a polyline; got:\n{source}"
    );
    let (circles, _, _) = marker_shape_counts(&source);
    assert_eq!(
        circles, 1,
        "the scatter series draws the circle it declares on its one point; got:\n{source}"
    );
}

/// The `(width, height, key-to-label gap)` of every legend key the source draws
/// as a filled box, in the order written.
///
/// A line series' key is a stroke and a marker rather than a box, so it is not
/// collected here — [`LEGEND_KEY_LEN_PT`] is what sizes that one.
fn emitted_legend_key_boxes(source: &str) -> Vec<(f64, f64, f64)> {
    source
        .lines()
        .filter_map(|line| {
            let after: &str = line.split_once("box[#box(width: ")?.1;
            // A line series' key opens the same way but takes a `baseline:`
            // after its height and paints its marker inside a nested block, so
            // the filled keys are the ones whose height is followed *directly*
            // by the colour. Merely containing a fill is not enough — the
            // nested marker carries one of its own.
            let height: &str = after.split_once("height: ")?.1;
            if !height.split_once("pt")?.1.starts_with(", fill: ") {
                return None;
            }
            Some((
                leading_pt(after)?,
                leading_pt(height)?,
                leading_pt(after.split_once("#h(")?.1)?,
            ))
        })
        .collect()
}

/// The #1169 workbook's bottom legend, restated in a face the crate measures
/// natively.
///
/// `Gift Budget and Tracker1.xlsx` (attached to #982) puts three bar series and
/// one overlaid line series in a bottom legend set in 9pt Segoe UI, which no
/// runner is guaranteed to have. Calibri's `hhea` metrics are compiled in — see
/// [`CALIBRI_CHART_LINE_METRICS_EM`] — so a legend restated in it is assertable
/// wherever the tests run.
fn excel_bottom_legend_chart(family: &str, size_pt: f64) -> Chart {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Column;
    chart.host = crate::ir::ChartHost::Spreadsheet;
    chart.categories = vec!["Jan".to_string()];
    chart.has_legend = true;
    chart.legend_position = LegendPosition::Bottom;
    chart.title = None;
    chart.auto_title_deleted = true;
    chart.text_font_family = Some(family.to_string());
    chart.text_style.size_pt = Some(size_pt);
    chart
}

#[test]
fn an_excel_bar_legend_key_is_the_flat_bar_excel_draws() {
    // Native Excel for Mac 16.112 exports of the #1169 workbook, one factor
    // changed per variant — five legend sizes and four faces against a
    // layout-identical re-zip control — draw a bar series' key as a wide flat
    // rectangle, never the 9pt square this renderer drew.
    //
    // Its width and the space before the label never move: 19.2000pt and
    // 2.0250pt in all ten exports, and 19.200pt again in the unrelated
    // `WithChart.xlsx` export, whose chart frame is under half as wide. Only
    // the height follows the legend's text, and it follows the face as well as
    // the size: 0.45 of the face's bare `hhea` line box.
    //
    // Calibri's box is 2500/2048 em, so 0.45 of it is 4.9438pt at 9pt against
    // the export's 4.9512 and 9.8877pt at 18pt against its 9.8902 — both inside
    // the 0.0122pt the GT's own 0.82 print scale can resolve.
    let measurements = [(9.0_f64, 4.9512_f64), (18.0, 9.8902)];

    for (size_pt, expected_height_pt) in measurements {
        let source: String = chart_source(excel_bottom_legend_chart("Calibri", size_pt));
        let keys: Vec<(f64, f64, f64)> = emitted_legend_key_boxes(&source);
        assert_eq!(
            keys.len(),
            2,
            "each of the two bar series takes a legend key; got:\n{source}"
        );
        for (width_pt, height_pt, gap_pt) in keys {
            assert!(
                (width_pt - 19.2).abs() <= 0.01,
                "{size_pt}pt: the key is {width_pt}pt wide, Excel's 19.2pt; got:\n{source}"
            );
            assert!(
                (height_pt - expected_height_pt).abs() <= 0.02,
                "{size_pt}pt: the key is {height_pt}pt tall, Excel's {expected_height_pt}pt; \
                 got:\n{source}"
            );
            assert!(
                (gap_pt - 2.025).abs() <= 0.005,
                "{size_pt}pt: the key leaves {gap_pt}pt before its label, Excel's 2.025pt; \
                 got:\n{source}"
            );
        }
    }
}

/// Fresh native Gift Budget exports vary only the legend size. Heights are
/// measured in printed points at the workbook's 0.82 fit scale (#1603).
#[test]
fn segoe_ui_legend_keys_retain_native_heights_without_the_office_font() {
    for (size, bold, native_height) in [
        (6.0, false, 2.9447),
        (9.0, false, 4.4169),
        (12.0, false, 5.8892),
        (18.0, false, 8.8357),
        (9.0, true, 4.4169),
    ] {
        let mut chart = excel_bottom_legend_chart("Segoe UI", size);
        chart.legend_text_style.bold = Some(bold);
        let source = chart_source(chart);
        let keys = emitted_legend_key_boxes(&source);
        assert_eq!(keys.len(), 2);
        for (width, height, _) in keys {
            assert!((width * 0.82 - 15.744).abs() < 0.01);
            assert!(
                (height * 0.82 - native_height).abs() < 0.01,
                "{size}pt Segoe UI: native key height {native_height}, got {}",
                height * 0.82
            );
        }
    }
}

/// Native after-export plot getters from the same workbook, with only its
/// value-axis size varied. The getter excludes the 4pt chart-area inset.
#[test]
fn segoe_ui_axis_gutters_preserve_native_plot_width_without_the_office_font() {
    let frame = (1015.977952755906, 307.973236083984);
    for (size, bold, inside_left, native_width) in [
        (6.0, false, 21.588503937008, 979.389448818898),
        (9.0, false, 31.121574803150, 969.856377952756),
        (12.0, false, 40.676614173228, 960.301338582677),
        (18.0, false, 59.767637795276, 941.210314960630),
        (9.0, true, 32.421732283465, 968.556220472441),
    ] {
        let mut chart = excel_bottom_legend_chart("Segoe UI", 9.0);
        chart.value_axis_text_style.size_pt = Some(size);
        chart.value_axis_text_style.bold = Some(bold);
        chart.value_axis_number_format = Some("\\$#,##0".to_string());
        chart.series[0].values = vec![180.0];
        let (left, _, right, _) = axis_plot_rect(&chart, frame, false);
        assert!(
            (left - (inside_left + 4.0)).abs() < 0.015,
            "{size}pt Segoe UI: native plot left {}, got {left}",
            inside_left + 4.0
        );
        assert!(
            (right - left - native_width).abs() < 0.015,
            "{size}pt Segoe UI: native plot width {native_width}, got {}",
            right - left
        );
    }
}

/// Pin the actual filled-key rectangles after parsing and compiling the public
/// workbook, including their baseline placement rather than only markup size.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn gift_budget_compiled_legend_keys_match_native_bounds() {
    let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1603_gift_budget.xlsx");
    let (doc, _) = crate::parser::Parser::parse(
        &crate::parser::xlsx::XlsxParser,
        data,
        &crate::config::ConvertOptions::default(),
    )
    .expect("the public gift-budget fixture parses");
    let output = generate_typst(&doc).unwrap();
    let pages = crate::render::pdf::compiled_page_paint_sequences(&output.source, &output.images)
        .expect("the gift-budget fixture compiles");
    assert_eq!(pages.len(), 2);
    let keys: Vec<_> = pages[1]
        .iter()
        .filter(|paint| {
            let (left, top, right, bottom) = paint.bounds;
            paint.rectangle_fill.is_some()
                && (right - left - 15.744).abs() < 0.01
                && top > 350.0
                && bottom < 365.0
        })
        .collect();
    let native = [
        (548.2121, 355.3613, 563.9561, 359.7782),
        (631.7659, 355.3613, 647.5099, 359.7782),
        (713.1958, 355.3613, 728.9398, 359.7782),
    ];
    assert_eq!(keys.len(), native.len());
    for (key, (left, top, right, bottom)) in keys.iter().zip(native) {
        let actual = key.bounds;
        assert!(
            (actual.0 - left).abs() < 0.5
                && (actual.1 - top).abs() < 0.5
                && (actual.2 - right).abs() < 0.5
                && (actual.3 - bottom).abs() < 0.5,
            "native key {:?}, actual {actual:?}",
            (left, top, right, bottom)
        );
        assert!(((actual.3 - actual.1) - (bottom - top)).abs() < 0.01);
    }
}

#[test]
fn an_excel_chartsheet_legend_key_scales_as_a_square_with_its_text() {
    // Excel for Mac 16.100 gives the `any_sheets.xlsx` chartsheet a native
    // 4.9433pt square key at 9pt, followed by a 2.0967pt gap. The shared
    // PowerPoint factors resolve to 4.9437pt and 2.096895pt respectively,
    // inside 0.0004pt of that export; unlike the 19.2pt flat key Excel uses
    // for an anchored worksheet chart, both dimensions scale with text
    // (#1315).
    for size_pt in [9.0_f64, 18.0] {
        let mut chart = excel_bottom_legend_chart("Calibri", size_pt);
        chart.host = crate::ir::ChartHost::SpreadsheetChartsheet;
        let source: String = chart_source(chart);
        let keys: Vec<(f64, f64, f64)> = emitted_legend_key_boxes(&source);
        assert_eq!(
            keys.len(),
            2,
            "each of the two bar series takes a legend key; got:\n{source}"
        );
        let expected_side_pt: f64 = PPTX_LEGEND_KEY_EM * size_pt;
        let expected_gap_pt: f64 =
            PPTX_LEGEND_KEY_LABEL_GAP_KEY_SHARE * expected_side_pt + PPTX_LEGEND_KEY_LABEL_GAP_PT;
        for (width_pt, height_pt, gap_pt) in keys {
            assert!(
                (width_pt - expected_side_pt).abs() <= 0.01
                    && (height_pt - expected_side_pt).abs() <= 0.01,
                "{size_pt}pt: a chartsheet key must be the {expected_side_pt}pt square Excel draws, got {width_pt}pt by {height_pt}pt in:\n{source}"
            );
            assert!(
                (gap_pt - expected_gap_pt).abs() <= 0.005,
                "{size_pt}pt: a chartsheet key leaves {expected_gap_pt}pt before its label, got {gap_pt}pt in:\n{source}"
            );
        }
    }
}

#[test]
fn an_excel_legend_key_keeps_one_length_across_both_families() {
    // The #1169 chart carries three bar series and one line laid over them, and
    // the export gives every one of the four keys the same 19.200pt span: the
    // bar keys as a filled rectangle, the line key as a stroke from 412.910 to
    // 432.110 in the `WithChart.xlsx` export. Sizing the two families apart
    // would put a 19.2pt bar key beside a 20pt line key in one legend row.
    let mut chart = excel_bottom_legend_chart("Calibri", 9.0);
    chart.series[1].plot_type = Some(ChartType::Line);
    let source: String = chart_source(chart);

    let keys: Vec<(f64, f64, f64)> = emitted_legend_key_boxes(&source);
    assert_eq!(
        keys.len(),
        1,
        "only the column series takes a filled key; got:\n{source}"
    );
    assert!(
        (keys[0].0 - LEGEND_KEY_LEN_PT).abs() <= 0.01,
        "the bar key spans {}pt where the line key spans {}pt; got:\n{source}",
        keys[0].0,
        format_f64(LEGEND_KEY_LEN_PT)
    );
    assert!(
        source.contains(&format!(
            "line(end: ({}pt, 0pt)",
            format_f64(LEGEND_KEY_LEN_PT)
        )),
        "the overlaid line's key samples the series across the same span; got:\n{source}"
    );
}

#[test]
fn a_powerpoint_legend_key_keeps_its_own_square() {
    // PowerPoint's automatic layout is a separate regime — its key is a square
    // scaled from chart text (#804) and was fitted to native 16.112 exports —
    // so the Excel measurement must not reach it.
    let mut chart = excel_bottom_legend_chart("Calibri", 12.0);
    chart.host = crate::ir::ChartHost::Presentation;
    let source: String = chart_source(chart);

    let keys: Vec<(f64, f64, f64)> = emitted_legend_key_boxes(&source);
    assert!(
        !keys.is_empty(),
        "the legend draws its keys; got:\n{source}"
    );
    for (width_pt, height_pt, _) in keys {
        assert!(
            (width_pt - height_pt).abs() < 1e-9,
            "a PowerPoint key stays square; got {width_pt}pt by {height_pt}pt in:\n{source}"
        );
        assert!(
            (width_pt - PPTX_LEGEND_KEY_EM * 12.0).abs() <= 0.01,
            "a PowerPoint key keeps its own text-scaled side; got {width_pt}pt in:\n{source}"
        );
    }
}

// ----- The plot rectangle c:plotArea/c:layout states (issue #1182) -----

/// The chart area the reported workbook's drawing anchor gives its
/// `january income:` bar chart, in points before the sheet's print scale:
/// `SheetChartPlacement { width: 307.98527559055117, height: 207.52251968503936 }`.
const MONTHLY_BUDGET_CHART_FRAME: (f64, f64) = (307.98527559055117, 207.52251968503936);

/// That chart's `c:plotArea/c:layout/c:manualLayout`, verbatim.
const MONTHLY_BUDGET_PLOT_LAYOUT: crate::ir::ChartPlotAreaLayout = crate::ir::ChartPlotAreaLayout {
    x: 0.3022229818508113,
    y: 0.34625485336714895,
    width: 0.6413931113556166,
    height: 0.5571170578930364,
};

/// The `january income:` bar chart of
/// `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`: five categories, one
/// series, no legend and no title. The template anchors it over the cells that
/// print `january income:` and `$1,225`, which is why its plot area starts a
/// third of the way down the chart.
fn monthly_budget_income_chart() -> Chart {
    let mut chart = bar_chart_at(
        None,
        &[
            "financial aid",
            "wages (after-tax)",
            "family help",
            "from savings",
            "other",
        ],
    );
    chart.series.truncate(1);
    chart.series[0].name = None;
    chart.series[0].values = vec![
        0.0,
        0.3673469387755102,
        0.16326530612244897,
        0.40816326530612246,
        0.061224489795918366,
    ];
    chart.has_legend = false;
    chart.host = crate::ir::ChartHost::Spreadsheet;
    chart.plot_area_layout = Some(MONTHLY_BUDGET_PLOT_LAYOUT);
    chart
}

/// A stated inner plot rectangle seats the plot where Excel seats it.
///
/// Measured on a native Excel for Mac 16 export of the workbook, staged and run
/// inside Excel's own sandbox container, traced with `mutool draw -F trace`.
/// The sheet prints at 0.78, so every figure below is the printed one divided
/// by that scale.
///
/// The chart area is 240.23 x 161.87pt printed, which is exactly this frame at
/// 0.78. Inside it the bars start at 152.75pt against a chart-area left edge of
/// 80.14pt, the value ticks run 0%-50% at a 30.82pt pitch — a 154.09pt plot —
/// the five category bands are 18.04pt apart for a 90.18pt plot, and the top
/// band's bar sits at 181.24pt. Undoing the scale gives (93.09, 71.87) for the
/// origin and 197.55 x 115.62pt for the extent, which is the frame times the
/// four fractions to within 0.03pt.
#[test]
fn a_stated_plot_rectangle_seats_the_plot_where_excel_seats_it() {
    let chart: Chart = monthly_budget_income_chart();

    let actual = axis_plot_rect(&chart, MONTHLY_BUDGET_CHART_FRAME, false);

    let expected = (93.09, 71.87, 290.65, 187.49);
    let errors = [
        ("left", actual.0, expected.0),
        ("top", actual.1, expected.1),
        ("right", actual.2, expected.2),
        ("bottom", actual.3, expected.3),
    ]
    .map(|(edge, actual, expected)| (edge, actual, expected, (actual - expected).abs()));
    assert!(
        errors.iter().all(|(_, _, _, error)| *error <= 0.1),
        "plot edges against the native export: {errors:?}"
    );
}

/// The fractions are of the chart area, not of one particular frame: the same
/// chart in a frame of another size lands on the same fractions of it.
///
/// Triangulation against a fix that hardcoded the reported chart's numbers.
#[test]
fn a_stated_plot_rectangle_scales_with_the_frame() {
    let chart: Chart = monthly_budget_income_chart();
    let frame: (f64, f64) = (480.0, 320.0);

    let (left, top, right, bottom) = axis_plot_rect(&chart, frame, false);

    let layout = MONTHLY_BUDGET_PLOT_LAYOUT;
    let expected = (
        layout.x * frame.0,
        layout.y * frame.1,
        (layout.x + layout.width) * frame.0,
        (layout.y + layout.height) * frame.1,
    );
    for (edge, actual, expected) in [
        ("left", left, expected.0),
        ("top", top, expected.1),
        ("right", right, expected.2),
        ("bottom", bottom, expected.3),
    ] {
        assert!(
            (actual - expected).abs() <= 0.01,
            "{edge} edge: {actual} against {expected}"
        );
    }
}

/// A chart that states no layout keeps the automatic one, which fills the
/// frame it was given less its chrome. Without this the fix would be free to
/// move every chart in the corpus.
#[test]
fn a_chart_without_a_stated_rectangle_keeps_the_automatic_plot() {
    let mut chart: Chart = monthly_budget_income_chart();
    chart.plot_area_layout = None;

    let (left, top, right, bottom) = axis_plot_rect(&chart, MONTHLY_BUDGET_CHART_FRAME, false);

    assert!(
        top < 1.0,
        "the automatic plot still starts at the frame's top edge, got {top}"
    );
    assert!(
        (bottom - (MONTHLY_BUDGET_CHART_FRAME.1 - chart_tick_band_pt(&chart))).abs() <= 0.01,
        "the automatic plot still ends above the tick band, got {bottom}"
    );
    assert!(
        (left - chart_category_gutter_pt(&chart)).abs() <= 0.01,
        "the automatic plot still starts right of the category gutter, got {left}"
    );
    assert!(
        (right - MONTHLY_BUDGET_CHART_FRAME.0).abs() <= 0.01,
        "the automatic plot still reaches the frame's right edge, got {right}"
    );
}

/// Everything the plot holds moves with it: the bars fill the stated rectangle
/// and the category labels stay against its left edge.
///
/// The bars and the labels are placed from the box's own edges rather than from
/// the plot's origin, so a displaced plot that left them behind would draw the
/// bars over the cells the chart floats above — which is what #1182 reports.
#[test]
fn the_bars_and_category_labels_follow_a_stated_plot_rectangle() {
    let chart: Chart = monthly_budget_income_chart();
    let (left, top, _, bottom) = axis_plot_rect(&chart, MONTHLY_BUDGET_CHART_FRAME, false);

    let source: String = framed_chart_source(
        &chart,
        MONTHLY_BUDGET_CHART_FRAME.0,
        MONTHLY_BUDGET_CHART_FRAME.1,
    );

    let bars: Vec<PlacedRect> = emitted_rects(&source);
    assert_eq!(bars.len(), chart.categories.len(), "one bar per category");
    for bar in &bars {
        assert!(
            (bar.dx - left).abs() <= 0.01,
            "every bar starts on the plot's left edge, got {bar:?}"
        );
        assert!(
            bar.dy >= top - 0.01 && bar.dy + bar.height <= bottom + 0.01,
            "every bar stands inside the plot, got {bar:?} for {top}..{bottom}"
        );
    }

    let label: PlacedBox = placed_box_holding(&source, "from savings");
    assert!(
        label.dx + label.width <= left + 0.01,
        "the category labels stay left of the plot, got {label:?}"
    );
    assert!(
        label.dy >= top - 0.01 && label.dy + label.height <= bottom + 0.01,
        "each category label sits beside its own band, got {label:?}"
    );
}

// ----- A stated plot rectangle that overruns its chart area (issue #1272) -----

/// The chart area the reported workbook's drawing anchor gives its
/// `january cash flow:` bar chart, in points before the sheet's 0.78 print
/// scale. The native export prints it 123.38 x 161.86pt, which is this at 0.78.
const CASH_FLOW_BAR_CHART_FRAME: (f64, f64) = (158.18, 207.51);

/// That chart's `c:plotArea/c:layout/c:manualLayout`, verbatim: a plot the
/// full chart-area width, starting 9.2% of the way in.
const CASH_FLOW_BAR_PLOT_LAYOUT: crate::ir::ChartPlotAreaLayout = crate::ir::ChartPlotAreaLayout {
    x: 0.091_961_072_570_217_49,
    y: 0.327_691_879_586_070_1,
    width: 1.0,
    height: 0.594_654_619_546_440_2,
};

/// One pull-back case: the layout as stated and the plot edges the native
/// export puts it at, as `[left, top, right, bottom]` fractions of the frame.
type PulledBackPlotCase = (&'static str, crate::ir::ChartPlotAreaLayout, [f64; 4]);

/// A stated rectangle that would run past the chart area is pulled back until
/// its far edge lands on the chart area's, along that axis alone.
///
/// Native Excel for Mac 16 exports of the reported workbook with one value of
/// the cash-flow chart's layout rewritten per variant, the plot read off its
/// plot-area fill and its category axis as fractions of the printed chart
/// area: the file's own `x` 0.092 with `w` 1 prints from the chart area's left
/// edge, and so do `x` 0.2 and `x` 0; `w` 0.95 starts at 0.05 and `x` 0.3 with
/// `w` 0.8 at 0.2, each given back exactly its overrun; `w` 0.9 keeps its 0.092
/// because it fits. Down the frame, `y` 0.5 comes back to 0.4053 and `h` 0.7
/// pulls `y` 0.3277 back to 0.3, the other axis holding still throughout.
/// Taking the fractions as written had put this chart's plot 13.9pt right of
/// the export's (issue #1272).
#[test]
fn a_stated_plot_rectangle_that_overruns_the_chart_area_is_pulled_back_inside_it() {
    let (frame_w, frame_h) = CASH_FLOW_BAR_CHART_FRAME;
    let as_written = CASH_FLOW_BAR_PLOT_LAYOUT;
    let cases: [PulledBackPlotCase; 8] = [
        (
            "x 0.092 w 1, as the file states",
            as_written,
            [0.0, 0.3277, 1.0, 0.9224],
        ),
        (
            "x 0.2 w 1",
            crate::ir::ChartPlotAreaLayout {
                x: 0.2,
                ..as_written
            },
            [0.0, 0.3277, 1.0, 0.9224],
        ),
        (
            "x 0 w 1",
            crate::ir::ChartPlotAreaLayout {
                x: 0.0,
                ..as_written
            },
            [0.0, 0.3277, 1.0, 0.9224],
        ),
        (
            "w 0.9 fits and stays where it is",
            crate::ir::ChartPlotAreaLayout {
                width: 0.9,
                ..as_written
            },
            [0.092, 0.3277, 0.992, 0.9224],
        ),
        (
            "w 0.95 gives back its 0.042 overrun",
            crate::ir::ChartPlotAreaLayout {
                width: 0.95,
                ..as_written
            },
            [0.05, 0.3277, 1.0, 0.9224],
        ),
        (
            "x 0.3 w 0.8 gives back its 0.1 overrun",
            crate::ir::ChartPlotAreaLayout {
                x: 0.3,
                width: 0.8,
                ..as_written
            },
            [0.2, 0.3277, 1.0, 0.9224],
        ),
        (
            "y 0.5 comes back up the frame",
            crate::ir::ChartPlotAreaLayout {
                y: 0.5,
                ..as_written
            },
            [0.0, 0.4053, 1.0, 1.0],
        ),
        (
            "h 0.7 pulls y back to 0.3",
            crate::ir::ChartPlotAreaLayout {
                height: 0.7,
                ..as_written
            },
            [0.0, 0.3, 1.0, 1.0],
        ),
    ];

    for (case, layout, [left, top, right, bottom]) in cases {
        let mut chart: Chart = cash_flow_bar_chart();
        chart.plot_area_layout = Some(layout);
        let actual = axis_plot_rect(&chart, CASH_FLOW_BAR_CHART_FRAME, false);
        let expected = (
            left * frame_w,
            top * frame_h,
            right * frame_w,
            bottom * frame_h,
        );
        let errors = [
            ("left", actual.0, expected.0),
            ("top", actual.1, expected.1),
            ("right", actual.2, expected.2),
            ("bottom", actual.3, expected.3),
        ]
        .map(|(edge, actual, expected)| (edge, actual, expected, (actual - expected).abs()));
        assert!(
            errors.iter().all(|(_, _, _, error)| *error <= 0.05),
            "{case}: plot edges {errors:?}"
        );
    }
}

// ----- A pulled-back bar plot keeps its value labels inside (issue #1634) -----

/// The chart area the reported workbook's drawing anchor gives its `january
/// income:` bar chart, in points before the sheet's 0.78 print scale. The
/// native export prints it 240.23pt wide, which is this at 0.78.
const INCOME_BAR_CHART_FRAME: (f64, f64) = (307.985, 207.523);

/// That chart's `c:plotArea/c:layout/c:manualLayout`, verbatim.
const INCOME_BAR_PLOT_LAYOUT: crate::ir::ChartPlotAreaLayout = crate::ir::ChartPlotAreaLayout {
    x: 0.302_222_981_850_811_3,
    y: 0.346_254_853_367_148_95,
    width: 0.641_393_111_355_616_6,
    height: 0.557_117_057_893_036_4,
};

/// The `january income:` chart: horizontal bars over five categories, the
/// value axis printing percentage tick labels under the plot from 0 to 0.5.
///
/// The labels are set in Calibri rather than the workbook's Trebuchet MS so the
/// expected overhang reads the same calibrated advance table on every runner.
fn income_bar_chart(value_format: &str, value_size_pt: f64) -> Chart {
    let mut chart: Chart = cash_flow_bar_chart();
    chart.host = crate::ir::ChartHost::Spreadsheet;
    chart.grouping = ChartGrouping::Clustered;
    chart.categories = [
        "other",
        "from savings",
        "family help",
        "wages (after-tax)",
        "financial aid",
    ]
    .map(str::to_string)
    .to_vec();
    chart.series.truncate(1);
    chart.series[0].values = vec![0.0, 0.0, 0.05, 0.4, 0.3];
    chart.value_axis_min = Some(0.0);
    chart.value_axis_max = Some(0.5);
    chart.value_axis_major_unit = Some(0.1);
    chart.value_axis_number_format = Some(value_format.to_string());
    chart.text_font_family = Some("Calibri".to_string());
    chart.text_style.size_pt = Some(10.0);
    chart.value_axis_text_style.size_pt = Some(value_size_pt);
    chart
}

/// Half the last value label's advance plus Excel's flat 6pt: how far inside
/// the chart area's right edge a pulled-back plot ends.
fn income_bar_overhang_pt(last_label: &str, value_size_pt: f64) -> f64 {
    chart_text_advance_em("Calibri", false, last_label).expect("Calibri is calibrated")
        * value_size_pt
        / 2.0
        + 6.0
}

/// Assert a plot's four edges, as `(left, top, right, bottom)` in points.
fn assert_plot_edges(case: &str, actual: (f64, f64, f64, f64), expected: (f64, f64, f64, f64)) {
    let errors = [
        ("left", actual.0, expected.0),
        ("top", actual.1, expected.1),
        ("right", actual.2, expected.2),
        ("bottom", actual.3, expected.3),
    ]
    .map(|(edge, actual, expected)| (edge, actual, expected, (actual - expected).abs()));
    assert!(
        errors.iter().all(|(_, _, _, error)| *error <= 0.01),
        "{case}: plot edges {errors:?}"
    );
}

/// A bar plot whose value tick labels run under it is pulled back until the
/// last label, centred on its right edge, stops 6pt inside the chart area —
/// not until the plot's own edge reaches the chart area's.
///
/// Native Excel for Mac 16.112 exports of the reported workbook with one value
/// of this chart's layout rewritten per variant end the plot at 0.9538 of the
/// chart area for `x` 0.33, 0.36, 0.40 and 0.50 alike — 0.33 would still have
/// ended inside it — and further in for a wider last label (`50.00%`, 0.9308)
/// or a larger one (`50%` at 14pt, 0.9431). The file's own `x` fits and stays.
/// Taking only the plot's own edge as the limit put the `x` 0.5 plot 11.1pt
/// right of the export (issue #1634).
#[test]
fn a_pulled_back_bar_plot_keeps_its_last_value_label_inside_the_chart_area() {
    let (frame_w, frame_h) = INCOME_BAR_CHART_FRAME;
    let layout = INCOME_BAR_PLOT_LAYOUT;
    let top: f64 = layout.y * frame_h;
    let bottom: f64 = top + layout.height * frame_h;
    let plot_w: f64 = layout.width * frame_w;

    let mut as_written: Chart = income_bar_chart("0%", 10.0);
    as_written.plot_area_layout = Some(layout);
    assert_plot_edges(
        "the file's own layout fits",
        axis_plot_rect(&as_written, INCOME_BAR_CHART_FRAME, false),
        (layout.x * frame_w, top, layout.x * frame_w + plot_w, bottom),
    );

    let cases: [(&str, f64, &str, &str, f64); 6] = [
        ("x 0.33", 0.33, "0%", "50%", 10.0),
        ("x 0.36", 0.36, "0%", "50%", 10.0),
        ("x 0.40", 0.40, "0%", "50%", 10.0),
        ("x 0.50", 0.50, "0%", "50%", 10.0),
        ("x 0.50, a wider last label", 0.50, "0.00%", "50.00%", 10.0),
        ("x 0.50, a larger last label", 0.50, "0%", "50%", 14.0),
    ];
    for (case, x, value_format, last_label, value_size_pt) in cases {
        let mut chart: Chart = income_bar_chart(value_format, value_size_pt);
        chart.plot_area_layout = Some(crate::ir::ChartPlotAreaLayout { x, ..layout });
        let right: f64 = frame_w - income_bar_overhang_pt(last_label, value_size_pt);
        assert_plot_edges(
            case,
            axis_plot_rect(&chart, INCOME_BAR_CHART_FRAME, false),
            (right - plot_w, top, right, bottom),
        );
    }
}

/// A stated bar plot too wide to pull back past its category labels is cut to
/// start where those labels end, rather than pushing them off the chart area.
///
/// Excel's export of `w` 0.9 on the reported chart runs the plot from 0.2863 to
/// 0.9538 of the chart area: the right edge at the labelled limit, the left
/// 1.525pt + 0.89152em past the widest category label, a band four exports at
/// 8, 10, 14 and 18pt category labels fit to 0.005pt. `x` 0.1 with `w` 0.9
/// prints on the same band, but `x` 0.1 with the file's own `w` fits and stays
/// at 0.1, the labels cut short to make room (issue #1634).
#[test]
fn a_bar_plot_too_wide_to_pull_back_is_cut_to_its_category_labels() {
    let (frame_w, frame_h) = INCOME_BAR_CHART_FRAME;
    let layout = INCOME_BAR_PLOT_LAYOUT;
    let top: f64 = layout.y * frame_h;
    let bottom: f64 = top + layout.height * frame_h;
    let right: f64 = frame_w - income_bar_overhang_pt("50%", 10.0);
    let widest_em: f64 = chart_text_advance_em("Calibri", false, "wages (after-tax)")
        .expect("Calibri is calibrated");
    let band_pt = |category_size_pt: f64| -> f64 {
        widest_em * category_size_pt + 1.525 + 0.89152 * category_size_pt
    };

    let cases: [(&str, f64, f64, f64); 3] = [
        ("w 0.9", layout.x, 10.0, band_pt(10.0)),
        ("x 0.1 w 0.9", 0.1, 10.0, band_pt(10.0)),
        ("w 0.9, 14pt category labels", layout.x, 14.0, band_pt(14.0)),
    ];
    for (case, x, category_size_pt, left) in cases {
        let mut chart: Chart = income_bar_chart("0%", 10.0);
        chart.category_axis_text_style.size_pt = Some(category_size_pt);
        chart.plot_area_layout = Some(crate::ir::ChartPlotAreaLayout {
            x,
            width: 0.9,
            ..layout
        });
        assert_plot_edges(
            case,
            axis_plot_rect(&chart, INCOME_BAR_CHART_FRAME, false),
            (left, top, right, bottom),
        );
    }

    assert!(
        band_pt(10.0) > 0.1 * frame_w,
        "the labels must need more than 0.1 of the frame for the fitting case \
         to test anything"
    );
    let mut fits: Chart = income_bar_chart("0%", 10.0);
    fits.plot_area_layout = Some(crate::ir::ChartPlotAreaLayout { x: 0.1, ..layout });
    assert_plot_edges(
        "x 0.1 fits",
        axis_plot_rect(&fits, INCOME_BAR_CHART_FRAME, false),
        (0.1 * frame_w, top, (0.1 + layout.width) * frame_w, bottom),
    );
}

/// The label reserve covers only what was measured: an Excel bar plot whose
/// value axis prints labels at a stated size. A column plot, a deleted value
/// axis, a PowerPoint chart and a chart stating no size keep #1272's limit on
/// the plot's own edge.
#[test]
fn the_value_label_reserve_stays_off_charts_it_was_not_measured_on() {
    let (frame_w, frame_h) = INCOME_BAR_CHART_FRAME;
    let layout = crate::ir::ChartPlotAreaLayout {
        x: 0.5,
        ..INCOME_BAR_PLOT_LAYOUT
    };
    let top: f64 = layout.y * frame_h;
    let bottom: f64 = top + layout.height * frame_h;
    let own_edge_limit = ((1.0 - layout.width) * frame_w, top, frame_w, bottom);

    let mut column: Chart = income_bar_chart("0%", 10.0);
    column.chart_type = ChartType::Column;
    let mut deleted_axis: Chart = income_bar_chart("0%", 10.0);
    deleted_axis.value_axis_deleted = true;
    let mut presentation: Chart = income_bar_chart("0%", 10.0);
    presentation.host = crate::ir::ChartHost::Presentation;
    let mut undeclared_size: Chart = income_bar_chart("0%", 10.0);
    undeclared_size.text_style.size_pt = None;
    undeclared_size.value_axis_text_style.size_pt = None;

    for (case, mut chart) in [
        ("column plot", column),
        ("deleted value axis", deleted_axis),
        ("PowerPoint host", presentation),
        ("no stated size", undeclared_size),
    ] {
        chart.plot_area_layout = Some(layout);
        assert_plot_edges(
            case,
            axis_plot_rect(&chart, INCOME_BAR_CHART_FRAME, false),
            own_edge_limit,
        );
    }
}

// ----- Line-family plot rectangles (issue #1265) -----

/// The chart area and inner plot rectangle of `xl/charts/chart2.xml` in
/// `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`, before the worksheet's
/// 0.78 print scale is applied.
const MONTHLY_BUDGET_LINE_CHART_FRAME: (f64, f64) = (786.71, 75.55);
const MONTHLY_BUDGET_LINE_PLOT_LAYOUT: crate::ir::ChartPlotAreaLayout =
    crate::ir::ChartPlotAreaLayout {
        x: 0.14276848195056766,
        y: 0.07430787730224597,
        width: 0.8572315180494323,
        height: 0.5596752035269608,
    };

/// A line/scatter plot must use the same stated inner rectangle as the bar
/// family instead of laying itself out from its private automatic gutters.
///
/// Native Excel seats this plot at 175.19..701.22pt on the printed page. Its
/// chart area starts at 87.59pt and the sheet prints at 0.78, giving the four
/// unscaled edges asserted below. The chart's real value axis is deleted; the
/// test enables only its axis line so all four plot edges can be read back
/// from generated chrome without changing the line/scatter layout path.
#[test]
fn a_line_family_plot_honours_its_stated_rectangle() {
    let mut chart: Chart = combo_line_and_scatter_chart();
    chart.plot_area_layout = Some(MONTHLY_BUDGET_LINE_PLOT_LAYOUT);
    chart.value_axis_deleted = false;

    let mut automatic_chart: Chart = chart.clone();
    automatic_chart.plot_area_layout = None;
    let automatic_source: String = framed_chart_source(
        &automatic_chart,
        MONTHLY_BUDGET_LINE_CHART_FRAME.0,
        MONTHLY_BUDGET_LINE_CHART_FRAME.1,
    );
    let automatic_jan: PlacedBox = placed_box_holding(&automatic_source, "jan");

    let source: String = framed_chart_source(
        &chart,
        MONTHLY_BUDGET_LINE_CHART_FRAME.0,
        MONTHLY_BUDGET_LINE_CHART_FRAME.1,
    );
    let (x, y, width, height) = plot_rect(&emitted_lines(&source));
    let layout = MONTHLY_BUDGET_LINE_PLOT_LAYOUT;
    let expected = (
        layout.x * MONTHLY_BUDGET_LINE_CHART_FRAME.0,
        layout.y * MONTHLY_BUDGET_LINE_CHART_FRAME.1,
        layout.width * MONTHLY_BUDGET_LINE_CHART_FRAME.0,
        layout.height * MONTHLY_BUDGET_LINE_CHART_FRAME.1,
    );

    for (edge, actual, expected) in [
        ("left", x, expected.0),
        ("top", y, expected.1),
        ("width", width, expected.2),
        ("height", height, expected.3),
    ] {
        assert!(
            (actual - expected).abs() <= 0.02,
            "line-family plot {edge}: expected {expected}, got {actual}\n{source}"
        );
    }

    let jan: PlacedBox = placed_box_holding(&source, "jan");
    let first_band_centre: f64 = expected.0 + expected.2 / (2.0 * chart.categories.len() as f64);
    assert!(
        (jan.dx + jan.width / 2.0 - first_band_centre).abs() <= 0.02,
        "the first category label must follow the stated plot; got {jan:?}, expected centre {first_band_centre}"
    );
    assert!(
        (jan.dy - automatic_jan.dy).abs() <= 0.02,
        "implicit-size labels retain their uncalibrated automatic band; got {jan:?}, automatic {automatic_jan:?}"
    );

    let value_tick: PlacedBox = placed_box_holding(&source, "200");
    assert!(
        (value_tick.dx + value_tick.width - (expected.0 - 6.0)).abs() <= 0.02,
        "the value labels must keep their automatic 6pt gap while following the stated plot; got {value_tick:?}"
    );
}

// ----- A value axis fixed to its own interval (issue #1184) -----

/// The `january cash flow:` chart of `xl/charts/chart1.xml` in the workbook of
/// #1123: a stacked bar over three categories whose value axis is fixed to
/// −400..400 in one major unit of 400, so value zero sits at the middle of the
/// plot rather than on its left edge.
///
/// Gridlines are suppressed so the only lines in the generated source are the
/// two axes. The parsed workbook's gridline presence is tested separately.
fn cash_flow_bar_chart() -> Chart {
    let mut chart = stacked_support_chart(ChartGrouping::Stacked);
    chart.chart_type = ChartType::Bar;
    chart.title = None;
    chart.has_legend = false;
    // Non-numeric so `emitted_axis_ticks` cannot mistake a category label for a
    // value tick; the workbook's own categories are the automatic 1, 2, 3.
    chart.categories = vec!["one".to_string(), "two".to_string(), "three".to_string()];
    chart.series.truncate(2);
    chart.series[0].name = Some("Positive".to_string());
    chart.series[0].values = vec![0.0, 169.0, 169.0];
    chart.series[1].name = Some("Negative".to_string());
    chart.series[1].values = vec![0.0, 0.0, -169.0];
    chart.value_axis_min = Some(-400.0);
    chart.value_axis_max = Some(400.0);
    chart.value_axis_major_unit = Some(400.0);
    chart.major_gridline_line = crate::ir::ChartLine::Suppressed;
    chart
}

/// The `dx`/`dy` a `#place` line puts its content at.
fn place_origin(line: &str) -> Option<(f64, f64)> {
    let dx: f64 = line
        .split("dx: ")
        .nth(1)?
        .split("pt")
        .next()?
        .trim()
        .parse()
        .ok()?;
    let dy: f64 = line
        .split("dy: ")
        .nth(1)?
        .split("pt")
        .next()?
        .trim()
        .parse()
        .ok()?;
    Some((dx, dy))
}

/// Origin and length of the one vertical axis line in the generated source.
fn vertical_axis_line(source: &str) -> (f64, f64, f64) {
    let line: &str = source
        .lines()
        .find(|line| line.contains("line(end: (0pt, "))
        .unwrap_or_else(|| panic!("no vertical axis line in:\n{source}"));
    let (dx, dy) = place_origin(line).expect("the axis line is placed");
    let length: f64 = line
        .split("line(end: (0pt, ")
        .nth(1)
        .and_then(|rest| rest.split("pt").next())
        .and_then(|value| value.parse().ok())
        .expect("the axis line states its length");
    (dx, dy, length)
}

/// Origin and length of the one horizontal axis line in the generated source.
fn horizontal_axis_line(source: &str) -> (f64, f64, f64) {
    let line: &str = source
        .lines()
        .find(|line| line.contains("pt, 0pt), stroke:"))
        .unwrap_or_else(|| panic!("no horizontal axis line in:\n{source}"));
    let (dx, dy) = place_origin(line).expect("the axis line is placed");
    let length: f64 = line
        .split("line(end: (")
        .nth(1)
        .and_then(|rest| rest.split("pt").next())
        .and_then(|value| value.parse().ok())
        .expect("the axis line states its length");
    (dx, dy, length)
}

/// Every drawn bar as `(dx, dy, width, height)`, skipping the empty rectangles
/// a zero-valued point still emits.
fn drawn_bars(source: &str) -> Vec<(f64, f64, f64, f64)> {
    source
        .lines()
        .filter(|line| line.contains("rect(width: "))
        .filter_map(|line| {
            let (dx, dy) = place_origin(line)?;
            let width: f64 = line
                .split("rect(width: ")
                .nth(1)?
                .split("pt")
                .next()?
                .parse()
                .ok()?;
            let height: f64 = line
                .split(", height: ")
                .nth(1)?
                .split("pt")
                .next()?
                .parse()
                .ok()?;
            Some((dx, dy, width, height))
        })
        .filter(|(_, _, width, height)| *width > 0.01 && *height > 0.01)
        .collect()
}

/// A stated `<c:min>`/`<c:max>` puts the category axis on the value-zero line
/// rather than on the plot's edge (issue #1184).
///
/// On a −400..400 axis zero is the middle of the plot, which is where Excel's
/// own export draws the axis of the `january cash flow:` chart.
#[test]
fn a_stated_axis_interval_seats_the_category_axis_on_value_zero() {
    let source: String = chart_source(cash_flow_bar_chart());

    // A bar chart's value axis runs along the bottom, so the horizontal line is
    // the value axis and gives the plot's own left edge and width.
    let (plot_x, _, plot_w) = horizontal_axis_line(&source);
    let (category_axis_x, _, _) = vertical_axis_line(&source);

    let expected: f64 = plot_x + plot_w / 2.0;
    assert!(
        (category_axis_x - expected).abs() < 0.01,
        "the category axis must stand on value zero, halfway across a −400..400 plot: \
         expected {expected}, got {category_axis_x} (plot {plot_x}..{})",
        plot_x + plot_w
    );
}

/// A negative segment draws on the far side of zero instead of collapsing onto
/// the plot floor (issue #1184).
#[test]
fn a_negative_stacked_segment_draws_on_the_far_side_of_zero() {
    let source: String = chart_source(cash_flow_bar_chart());

    let (plot_x, _, plot_w) = horizontal_axis_line(&source);
    let zero_x: f64 = plot_x + plot_w / 2.0;
    // 169 of the 800 the axis spans.
    let expected_w: f64 = 169.0 / 800.0 * plot_w;

    let bars: Vec<(f64, f64, f64, f64)> = drawn_bars(&source);
    assert_eq!(
        bars.len(),
        3,
        "two positive segments and one negative one are drawn, got {bars:?}"
    );
    for (dx, _, width, _) in &bars {
        assert!(
            (width - expected_w).abs() < 0.01,
            "every 169-unit segment spans the same share of the axis: \
             expected {expected_w}, got {width} at dx {dx}"
        );
    }

    let positives: usize = bars
        .iter()
        .filter(|(dx, _, _, _)| (dx - zero_x).abs() < 0.01)
        .count();
    assert_eq!(
        positives, 2,
        "both +169 segments start on the zero line, got {bars:?}"
    );

    let negative: &(f64, f64, f64, f64) = bars
        .iter()
        .find(|(dx, _, _, _)| *dx < zero_x - 0.01)
        .unwrap_or_else(|| panic!("the −169 segment must reach left of zero, got {bars:?}"));
    assert!(
        (negative.0 + negative.2 - zero_x).abs() < 0.01,
        "the −169 segment must end on the zero line, got {negative:?} against zero at {zero_x}"
    );
}

/// The ticks run from the stated minimum in the stated unit, so a −400..400
/// axis at 400 a unit is labelled −400, 0, 400 (issue #1184).
#[test]
fn a_stated_axis_interval_ticks_from_its_own_minimum() {
    let ticks: Vec<f64> = emitted_axis_ticks(&chart_source(cash_flow_bar_chart()));

    assert_eq!(ticks, vec![-400.0, 0.0, 400.0], "got {ticks:?}");
}

/// A stated interval with no `<c:majorUnit>` takes its tick interval from that
/// interval rather than from the data behind it (issue #1184).
///
/// `xl/charts/chart3.xml` of the workbook in #1123 fixes its value axis to
/// 0..0.5 over a 0.408 maximum, and the native Excel export ticks 0%, 10%, 20%,
/// 30%, 40%, 50%. Sizing the unit from the data gives 5% and twice as many
/// ticks.
#[test]
fn a_stated_interval_takes_its_unit_from_the_interval() {
    let mut chart = stacked_support_chart(ChartGrouping::Clustered);
    chart.title = None;
    chart.has_legend = false;
    chart.series.truncate(1);
    chart.series[0].values = vec![0.408, 0.367, 0.163];
    chart.value_axis_min = Some(0.0);
    chart.value_axis_max = Some(0.5);
    chart.major_gridline_line = crate::ir::ChartLine::Suppressed;

    let ticks: Vec<f64> = emitted_axis_ticks(&chart_source(chart));

    assert_eq!(
        ticks.len(),
        6,
        "0..0.5 in tenths is six ticks, got {ticks:?}"
    );
    assert!(
        (ticks[1] - 0.1).abs() < 1e-9 && (ticks[5] - 0.5).abs() < 1e-9,
        "got {ticks:?}"
    );
}

/// Data that reaches below zero pulls the automatic minimum down with it, which
/// is what lets a line dip under its category axis (issue #1184).
///
/// `xl/charts/chart2.xml` of the same workbook states only its maximum, so the
/// floor its June (−771) and September (−721) points need has to come from the
/// automatic scale.
#[test]
fn automatic_scaling_reaches_below_zero_for_negative_data() {
    let mut chart = stacked_support_chart(ChartGrouping::Clustered);
    chart.title = None;
    chart.has_legend = false;
    chart.series.truncate(1);
    chart.series[0].values = vec![109.0, -771.0, 34.0];
    chart.major_gridline_line = crate::ir::ChartLine::Suppressed;

    let source: String = chart_source(chart);
    let ticks: Vec<f64> = emitted_axis_ticks(&source);
    let axis_min: f64 = ticks.iter().copied().fold(f64::INFINITY, f64::min);
    assert!(
        axis_min <= -771.0,
        "the axis has to reach the −771 point, got {ticks:?}"
    );

    // A column chart's value axis is the vertical one, so the horizontal line
    // is the category axis and it must have left the plot floor.
    let (_, plot_y, plot_h) = vertical_axis_line(&source);
    let (_, category_axis_y, _) = horizontal_axis_line(&source);
    assert!(
        category_axis_y > plot_y + 0.01 && category_axis_y < plot_y + plot_h - 0.01,
        "the category axis must stand on zero inside the plot: got {category_axis_y} \
         in {plot_y}..{}",
        plot_y + plot_h
    );

    // The −771 bar hangs from the axis rather than growing up from the floor.
    let bars: Vec<(f64, f64, f64, f64)> = drawn_bars(&source);
    assert!(
        bars.iter()
            .any(|(_, dy, _, height)| (dy - category_axis_y).abs() < 0.01 && *height > 1.0),
        "the −771 column must start on the axis and hang below it, got {bars:?} \
         against an axis at {category_axis_y}"
    );
}

/// Positive-only data keeps the axis it always had: zero on the plot floor and
/// the category axis drawn there (issue #1184 must not move an ordinary chart).
#[test]
fn positive_only_data_keeps_the_category_axis_on_the_plot_floor() {
    let mut chart = stacked_support_chart(ChartGrouping::Stacked);
    chart.major_gridline_line = crate::ir::ChartLine::Suppressed;
    let source: String = chart_source(chart);

    let (_, plot_y, plot_h) = vertical_axis_line(&source);
    let (_, category_axis_y, _) = horizontal_axis_line(&source);

    assert!(
        (category_axis_y - (plot_y + plot_h)).abs() < 0.01,
        "an all-positive chart draws its category axis on the plot floor: \
         expected {}, got {category_axis_y}",
        plot_y + plot_h
    );
}

/// A line point below zero dips under the category axis instead of flattening
/// onto the plot floor (issue #1184).
#[test]
fn a_line_point_below_zero_dips_under_the_category_axis() {
    let mut chart = stacked_support_chart(ChartGrouping::Clustered);
    chart.chart_type = ChartType::Line;
    chart.title = None;
    chart.has_legend = false;
    chart.categories = vec![
        "jun".to_string(),
        "jul".to_string(),
        "aug".to_string(),
        "sep".to_string(),
    ];
    chart.series.truncate(1);
    chart.series[0].values = vec![-771.0, 109.0, 34.0, -721.0];
    chart.value_axis_max = Some(1000.0);
    chart.value_axis_major_unit = Some(200.0);
    chart.major_gridline_line = crate::ir::ChartLine::Suppressed;

    let source: String = chart_source(chart);
    let (_, category_axis_y, _) = horizontal_axis_line(&source);

    let path_line: &str = source
        .lines()
        .find(|line| line.contains("curve(stroke:"))
        .unwrap_or_else(|| panic!("no series polyline in:\n{source}"));
    let point_ys: Vec<f64> = path_line
        .split("pt, ")
        .skip(1)
        .filter_map(|chunk| chunk.split("pt)").next()?.trim().parse::<f64>().ok())
        .collect();
    assert_eq!(point_ys.len(), 4, "four points, got {point_ys:?} ");

    let lowest: f64 = point_ys.iter().copied().fold(f64::NEG_INFINITY, f64::max);
    assert!(
        lowest > category_axis_y + 1.0,
        "a −771 point must sit below the category axis at {category_axis_y}, got {point_ys:?}"
    );
    // The two positive points stay above it.
    assert_eq!(
        point_ys
            .iter()
            .filter(|y| **y < category_axis_y - 1.0)
            .count(),
        2,
        "the +109 and +34 points stay above the axis, got {point_ys:?}"
    );
}

// ----- The shapes a chart lays over itself (issue #1186) -----

/// The chart area the reported workbook's drawing anchor gives its cash-flow
/// line chart, in points before the sheet's 0.78 print scale:
/// `SheetChartPlacement { width: 786.7143307086615, height: 75.54897637795276 }`.
/// The native export prints it 613.63 x 58.93pt, which is exactly that at 0.78.
const CASH_FLOW_CHART_FRAME: (f64, f64) = (786.7143307086615, 75.54897637795276);

/// The `CASH FLOW` caption `xl/drawings/drawing2.xml` anchors over that chart,
/// verbatim: the whole left edge of the chart area, an eighth of its width, and
/// a third of its height down from the top.
fn cash_flow_caption() -> crate::ir::ChartUserShape {
    crate::ir::ChartUserShape {
        from: (0.0, 0.17913),
        extent: crate::ir::ChartUserShapeExtent::Corner {
            x: 0.11685,
            y: 0.50958,
        },
        paragraphs: vec![Paragraph {
            style: ParagraphStyle {
                space_before: Some(0.0),
                space_after: Some(0.0),
                ..ParagraphStyle::default()
            },
            runs: vec![Run {
                text: "CASH FLOW".to_string(),
                style: TextStyle {
                    font_family: Some("Cambria".to_string()),
                    font_size: Some(15.0),
                    bold: Some(true),
                    color: Some(Color::new(0x24, 0x67, 0x78)),
                    ..TextStyle::default()
                },
                href: None,
                footnote: None,
            }],
        }],
        text_insets: crate::ir::Insets {
            top: 3.6,
            right: 7.2,
            bottom: 3.6,
            left: 7.2,
        },
        fill: None,
        border: None,
        no_wrap: true,
    }
}

fn cash_flow_chart_with_caption() -> Chart {
    let mut chart: Chart = combo_line_and_scatter_chart();
    chart.host = crate::ir::ChartHost::Spreadsheet;
    chart.user_shapes = vec![cash_flow_caption()];
    chart
}

/// The line the generator wrote for the caption.
fn caption_markup(source: &str) -> String {
    source
        .lines()
        .find(|line| line.contains("[CASH FLOW]"))
        .unwrap_or_else(|| panic!("nothing draws the caption in:\n{source}"))
        .to_string()
}

/// A user shape's box is its anchor's fractions of the chart area, with the
/// corner's offset down the area rounded to a whole point — which is what the
/// native export moves the caption by when `cdr:from/cdr:y` is rewritten:
/// 0.17913 -> 0.25 -> 0.3 shifts it 5.00 then 9.00pt, the steps between 14, 19
/// and 23, not the 5.35 and 9.13 of the unrounded fractions.
#[test]
fn a_user_shape_takes_its_box_from_the_chart_area_fractions() {
    let chart: Chart = cash_flow_chart_with_caption();
    let source: String =
        framed_chart_source(&chart, CASH_FLOW_CHART_FRAME.0, CASH_FLOW_CHART_FRAME.1);
    let caption: PlacedBox = placed_box_holding(&source, "CASH FLOW");

    assert_eq!(
        caption.dx, 0.0,
        "the caption starts on the chart's left edge"
    );
    assert_eq!(
        caption.dy,
        (0.17913 * CASH_FLOW_CHART_FRAME.1).round(),
        "the corner's offset down the area is a whole point, got {caption:?}"
    );
    assert!(
        (caption.width - 0.11685 * CASH_FLOW_CHART_FRAME.0).abs() < 0.01,
        "the box spans `cdr:to` minus `cdr:from` across the area, got {caption:?}"
    );
    assert!(
        (caption.height - (0.50958 - 0.17913) * CASH_FLOW_CHART_FRAME.1).abs() < 0.01,
        "and the same down it, got {caption:?}"
    );
}

/// The text's own seat inside that box: the left inset exactly, the top inset
/// rounded to a whole point, and a first baseline a whole-point ascent below
/// the line's top edge.
///
/// Every term is one factor of the native export. Rewriting `lIns="0"` moves
/// the caption 7.20pt left; `tIns="0"` moves it 4.00pt up, which is the whole
/// point 3.6 rounds to; and sweeping `a:rPr@sz` over 10/15/20/30/40pt moves the
/// baseline by -4.00/0/+5.00/+15.00/+24.00pt, the steps of `round(0.9502 x sz)`
/// for the 0.95020em ascent Cambria is measured at.
#[test]
fn a_user_shape_seats_its_text_where_the_native_export_seats_it() {
    let chart: Chart = cash_flow_chart_with_caption();
    let source: String =
        framed_chart_source(&chart, CASH_FLOW_CHART_FRAME.0, CASH_FLOW_CHART_FRAME.1);
    let caption: String = caption_markup(&source);

    let inner: &str = caption
        .split_once("pt)[")
        .expect("the caption's box opens")
        .1;
    assert!(
        inner.contains("dx: 7.2pt"),
        "the pen starts at the left inset, unrounded: {caption}"
    );
    assert!(
        inner.contains("dy: 4pt"),
        "the top inset is a whole point: {caption}"
    );

    let Some((ascent_em, _)) = chart_face_line_metrics_em("Cambria", true) else {
        return; // no font search: wasm resolves no face to measure
    };
    let above_pt: f64 = (ascent_em * 15.0).round();
    assert!(
        caption.contains(&format!("top-edge: {}pt", format_f64(above_pt))),
        "the first baseline sits a whole-point ascent below the text's top: {caption}"
    );
    if (ascent_em - 0.9502).abs() < 0.0005 {
        // Cambria itself resolved, so the three terms are the export's own:
        // 14 + 4 + 14 = 32pt from the chart area's top to the baseline.
        assert_eq!(
            (0.17913 * CASH_FLOW_CHART_FRAME.1).round() + 4.0 + above_pt,
            32.0
        );
    }
}

/// `<a:bodyPr wrap="none"/>` lets the line run past the shape's own width: the
/// caption is 80pt of glyphs in a 91.93pt box whose insets leave 77.53pt, so a
/// bounded body would break `CASH FLOW` across two lines. A body that does not
/// state it stays inside those insets.
#[test]
fn a_wrap_none_body_is_not_bounded_by_the_shape_it_sits_in() {
    let mut chart: Chart = cash_flow_chart_with_caption();
    let unbounded: String = caption_markup(&framed_chart_source(
        &chart,
        CASH_FLOW_CHART_FRAME.0,
        CASH_FLOW_CHART_FRAME.1,
    ));
    assert!(
        unbounded.contains("box(width: auto)"),
        "a `wrap=\"none\"` body takes no width: {unbounded}"
    );

    chart.user_shapes[0].no_wrap = false;
    let bounded: String = caption_markup(&framed_chart_source(
        &chart,
        CASH_FLOW_CHART_FRAME.0,
        CASH_FLOW_CHART_FRAME.1,
    ));
    let inner_w: f64 = 0.11685 * CASH_FLOW_CHART_FRAME.0 - 7.2 - 7.2;
    assert!(
        bounded.contains(&format!("box(width: {}pt)", format_f64(inner_w))),
        "a wrapping body is bounded by what the insets leave: {bounded}"
    );
}

/// The run's own properties reach the page: Excel sets this caption in 15pt
/// bold Cambria at `#246778`, the theme's `accent1` at half the luminance.
#[test]
fn a_user_shape_run_keeps_its_size_weight_face_and_colour() {
    let chart: Chart = cash_flow_chart_with_caption();
    let source: String =
        framed_chart_source(&chart, CASH_FLOW_CHART_FRAME.0, CASH_FLOW_CHART_FRAME.1);
    let caption: String = caption_markup(&source);

    for expected in [
        "size: 15pt",
        "weight: \"bold\"",
        "fill: rgb(36, 103, 120)",
        "Cambria",
    ] {
        assert!(
            caption.contains(expected),
            "the caption should carry {expected}: {caption}"
        );
    }
}

/// A chart naming no user shapes emits nothing for them, so no existing output
/// moves.
#[test]
fn a_chart_with_no_user_shapes_draws_none() {
    let chart: Chart = combo_line_and_scatter_chart();
    let source: String =
        framed_chart_source(&chart, CASH_FLOW_CHART_FRAME.0, CASH_FLOW_CHART_FRAME.1);
    assert!(!source.contains("top-edge"), "no shape body: {source}");
}

#[test]
fn blank_category_number_format_hides_budget_axis_labels_without_moving_the_plot() {
    use std::io::Read;
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
    let mut archive = zip::ZipArchive::new(std::fs::File::open(fixture).unwrap()).unwrap();
    let mut xml = String::new();
    archive
        .by_name("xl/charts/chart1.xml")
        .unwrap()
        .read_to_string(&mut xml)
        .unwrap();
    let colors = std::collections::HashMap::new();
    let aliases = std::collections::HashMap::new();
    let scheme = crate::parser::drawingml::SchemeColors {
        colors: &colors,
        aliases: &aliases,
    };
    let chart = crate::parser::chart::parse_chart_xml(&xml, &scheme).unwrap();
    let visible = crate::parser::chart::parse_chart_xml(
        &xml.replace("formatCode=\";;\"", "formatCode=\"General\""),
        &scheme,
    )
    .unwrap();
    assert_eq!(chart.categories, ["1", "2", "3"]);
    let source = framed_chart_source(&chart, 220.0, 200.0);
    let control = framed_chart_source(&visible, 220.0, 200.0);
    let is_category = |line: &&str| {
        ["[1]", "[2]", "[3]"]
            .iter()
            .any(|label| line.contains(label))
    };
    assert_eq!(control.lines().filter(is_category).count(), 3);
    assert_eq!(
        source.lines().filter(is_category).count(),
        0,
        "blank catAx numFmt must emit no labels"
    );
    let without_labels = control
        .lines()
        .filter(|line| !is_category(line))
        .collect::<Vec<_>>();
    assert_eq!(
        source.lines().collect::<Vec<_>>(),
        without_labels,
        "blank labels keep every axis and plotted mark unchanged"
    );
}

#[test]
fn category_number_formats_apply_across_axis_chart_families() {
    for chart_type in [
        ChartType::Bar,
        ChartType::Column,
        ChartType::Line,
        ChartType::Area,
        ChartType::Other(crate::ir::RADAR_CHART_LABEL.to_string()),
    ] {
        let mut chart = sized_bar_chart(11.0);
        chart.chart_type = chart_type.clone();
        chart.categories = ["12.5", "-7.5", "0"].map(str::to_string).to_vec();
        chart.series[0].values = vec![100.0, 200.0, 300.0];
        chart.value_axis_deleted = true;
        let control = framed_chart_source(&chart, 480.0, 320.0);
        for format in [";;", ";;;", "\"\";\"\";\"\""] {
            chart.category_axis_number_format = Some(format.to_string());
            let source = framed_chart_source(&chart, 480.0, 320.0);
            let is_category = |line: &&str| {
                ["[12.5]", "[\\-7.5]", "[0]"]
                    .iter()
                    .any(|label| line.contains(label))
            };
            assert_eq!(
                control.lines().filter(is_category).count(),
                3,
                "{chart_type:?}"
            );
            assert_eq!(
                source.lines().filter(is_category).count(),
                0,
                "{chart_type:?} {format}"
            );
            assert_eq!(
                source.lines().collect::<Vec<_>>(),
                control
                    .lines()
                    .filter(|line| !is_category(line))
                    .collect::<Vec<_>>(),
                "{chart_type:?}: blanking must preserve plot geometry"
            );
        }
        chart.category_axis_number_format = Some("0.0".to_string());
        let formatted = framed_chart_source(&chart, 480.0, 320.0);
        for label in ["[12.5]", "[\\-7.5]", "[0.0]"] {
            assert!(formatted.contains(label), "{chart_type:?}: {label}");
        }
    }
}

#[test]
fn category_number_formats_keep_text_labels_and_quoted_semicolons() {
    let mut chart = sized_bar_chart(11.0);
    chart.categories = vec!["Tuition".to_string(), "12.5".to_string()];
    chart.category_axis_number_format = Some(";;".to_string());
    let source = framed_chart_source(&chart, 480.0, 320.0);
    assert!(source.contains("[Tuition]"));
    assert!(!source.contains("[12.5]"));
    chart.category_axis_number_format = Some("\"term;first\"".to_string());
    let source = framed_chart_source(&chart, 480.0, 320.0);
    assert!(source.contains("[term;first]"));
    assert!(source.contains("[Tuition]"));
}

#[test]
fn the_budget_cash_flow_chart_paints_only_its_explicitly_filled_bar() {
    use std::io::Read;
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
    let mut archive = zip::ZipArchive::new(std::fs::File::open(fixture).unwrap()).unwrap();
    let mut xml = String::new();
    archive
        .by_name("xl/charts/chart1.xml")
        .unwrap()
        .read_to_string(&mut xml)
        .unwrap();
    let colors =
        std::collections::HashMap::from([("accent3".to_string(), Color::new(0x48, 0xbb, 0x88))]);
    let aliases = std::collections::HashMap::new();
    let scheme = crate::parser::drawingml::SchemeColors {
        colors: &colors,
        aliases: &aliases,
    };
    let chart = crate::parser::chart::parse_chart_xml(&xml, &scheme).unwrap();
    let source = framed_chart_source(&chart, 220.0, 200.0);
    let painted_bars: Vec<_> = source
        .lines()
        .filter(|line| {
            line.contains("rect(width:")
                && line.contains("stroke: none))")
                && !line.contains("width: 0pt")
                && !line.contains("height: 0pt")
                && !line.contains("fill: none")
        })
        .collect();
    assert_eq!(
        painted_bars.len(),
        1,
        "series and point noFill must leave only the explicit green bar: {painted_bars:?}"
    );
}

#[test]
fn no_fill_overrides_preserve_filled_chart_geometry_and_data_labels() {
    use crate::ir::ChartFillMode;
    for family in [
        ChartType::Bar,
        ChartType::Column,
        ChartType::Pie,
        ChartType::Doughnut,
    ] {
        let mut chart = sized_bar_chart(11.0);
        chart.chart_type = family.clone();
        chart.has_legend = false;
        chart.categories = ["First", "Second", "Third"].map(str::to_string).to_vec();
        chart.series[0].values = vec![100.0, 200.0, 300.0];
        chart.series[0].fill = Some(Color::new(18, 52, 86));
        chart.series[0].data_labels.show_value = true;
        let before = framed_chart_source(&chart, 480.0, 320.0);
        chart.series[0].fill_mode = ChartFillMode::Suppressed;
        chart.series[0].point_fill_modes = vec![
            ChartFillMode::Automatic,
            ChartFillMode::Explicit,
            ChartFillMode::Suppressed,
        ];
        let after = framed_chart_source(&chart, 480.0, 320.0);
        assert_eq!(
            before.matches("fill: rgb(18, 52, 86)").count(),
            3,
            "{family:?}"
        );
        assert_eq!(
            after.matches("fill: rgb(18, 52, 86)").count(),
            1,
            "{family:?}"
        );
        let without_fill = |source: &str| {
            source
                .replace("fill: rgb(18, 52, 86)", "fill: PAINT")
                .replace("fill: none", "fill: PAINT")
        };
        assert_eq!(
            without_fill(&before),
            without_fill(&after),
            "{family:?}: only paint changes"
        );
    }
}

#[test]
fn no_fill_keeps_positive_and_negative_segments_in_the_stack() {
    use crate::ir::ChartFillMode;
    for family in [ChartType::Bar, ChartType::Column] {
        let mut chart = sized_bar_chart(11.0);
        chart.chart_type = family;
        chart.grouping = ChartGrouping::Stacked;
        chart.has_legend = false;
        chart.categories = ["Q1", "Q2", "Q3"].map(str::to_string).to_vec();
        chart.series[0].values = vec![-10.0, 20.0, 30.0];
        chart.series[0].fill = Some(Color::new(18, 52, 86));
        chart.series[0].data_labels.show_value = true;
        let mut second = chart.series[0].clone();
        second.values = vec![5.0, -15.0, 10.0];
        second.fill = Some(Color::new(98, 76, 54));
        chart.series.push(second);
        let before = framed_chart_source(&chart, 480.0, 320.0);
        chart.series[0].fill_mode = ChartFillMode::Suppressed;
        let after = framed_chart_source(&chart, 480.0, 320.0);
        assert_eq!(before.matches("fill: rgb(18, 52, 86)").count(), 3);
        assert_eq!(before.replace("fill: rgb(18, 52, 86)", "fill: none"), after);
    }
}

#[test]
fn no_fill_on_a_series_does_not_suppress_its_line_or_markers() {
    for family in [ChartType::Line, ChartType::Area, ChartType::Scatter] {
        let mut chart = sized_bar_chart(11.0);
        chart.chart_type = family;
        chart.series[0].fill = Some(Color::new(18, 52, 86));
        let before = framed_chart_source(&chart, 480.0, 320.0);
        chart.series[0].fill_mode = crate::ir::ChartFillMode::Suppressed;
        let after = framed_chart_source(&chart, 480.0, 320.0);
        assert_eq!(
            before, after,
            "shape noFill does not change a series stroke or marker"
        );
    }
}

#[test]
fn absent_major_gridlines_do_not_paint_rules_in_budget_charts() {
    use std::io::Read;
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
    let mut archive = zip::ZipArchive::new(std::fs::File::open(fixture).unwrap()).unwrap();
    for part in ["xl/charts/chart1.xml", "xl/charts/chart2.xml"] {
        let mut xml = String::new();
        archive
            .by_name(part)
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        assert!(!xml.contains("majorGridlines"));
        let colors = std::collections::HashMap::new();
        let aliases = std::collections::HashMap::new();
        let scheme = crate::parser::drawingml::SchemeColors {
            colors: &colors,
            aliases: &aliases,
        };
        let mut chart = crate::parser::chart::parse_chart_xml(&xml, &scheme).unwrap();
        let actual = framed_chart_source(&chart, 480.0, 240.0);
        chart.major_gridline_line = crate::ir::ChartLine::Suppressed;
        let without_rules = framed_chart_source(&chart, 480.0, 240.0);
        assert!(
            actual == without_rules,
            "{part}: absent majorGridlines must not add rules"
        );
    }
}

#[test]
fn budget_gridlines_keep_their_declared_opacity() {
    use std::io::Read;
    let fixture = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
    let mut archive = zip::ZipArchive::new(std::fs::File::open(fixture).unwrap()).unwrap();
    let colors = std::collections::HashMap::from([("bg1".to_string(), Color::new(255, 255, 255))]);
    let aliases = std::collections::HashMap::new();
    let scheme = crate::parser::drawingml::SchemeColors {
        colors: &colors,
        aliases: &aliases,
    };
    for part in ["xl/charts/chart3.xml", "xl/charts/chart4.xml"] {
        let mut xml = String::new();
        archive
            .by_name(part)
            .unwrap()
            .read_to_string(&mut xml)
            .unwrap();
        assert!(xml.contains(r#"<a:alpha val="25000"/>"#));
        let opaque_xml = xml.replace(r#"<a:alpha val="25000"/>"#, "");
        let opaque_chart = crate::parser::chart::parse_chart_xml(&opaque_xml, &scheme).unwrap();
        let opaque_source = framed_chart_source(&opaque_chart, 480.0, 240.0);
        for (alpha, byte) in [(25000, 64), (50000, 128), (0, 0), (100000, 255)] {
            let variant = xml.replace(
                r#"<a:alpha val="25000"/>"#,
                &format!(r#"<a:alpha val="{alpha}"/>"#),
            );
            let chart = crate::parser::chart::parse_chart_xml(&variant, &scheme).unwrap();
            let source = framed_chart_source(&chart, 480.0, 240.0);
            assert!(
                source.contains(&format!("0.25pt + rgb(191, 191, 191, {byte})")),
                "{part}: gridline alpha {alpha} must survive parsing and rendering"
            );
            assert_eq!(
                source.replace(&format!("rgb(191, 191, 191, {byte})"), "rgb(191, 191, 191)"),
                opaque_source,
                "{part}: opacity must not change chart geometry or labels"
            );
        }
    }
}

#[test]
fn budget_month_axis_uses_its_major_theme_font() {
    let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
    let (doc, _) = crate::parser::Parser::parse(
        &crate::parser::xlsx::XlsxParser,
        data,
        &crate::config::ConvertOptions::default(),
    )
    .expect("the budget fixture parses");
    let source = generate_typst(&doc).unwrap().source;
    for label in [
        "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec", "year",
    ] {
        // Source categories retain trailing spaces (including escaped runs).
        let prefix = format!(")[{label}");
        let lines: Vec<_> = source
            .lines()
            .filter(|line| line.contains("#place(") && line.contains(&prefix))
            .collect();
        assert_eq!(
            lines.len(),
            1,
            "one category label for {label}: {}",
            source
                .lines()
                .filter(|line| line.contains(label))
                .collect::<Vec<_>>()
                .join("\n")
        );
        assert!(
            lines[0].contains("font:") && lines[0].contains("Cambria"),
            "axis +mj-lt must override the chart's Trebuchet MS: {}",
            lines[0]
        );
    }
}

#[test]
fn axis_typefaces_reach_each_axis_chart_renderer() {
    for kind in [
        ChartType::Bar,
        ChartType::Column,
        ChartType::Line,
        ChartType::Area,
        ChartType::Other(crate::ir::RADAR_CHART_LABEL.to_string()),
    ] {
        let mut chart = sized_bar_chart(11.0);
        chart.chart_type = kind.clone();
        chart.categories = vec!["Q1".to_string(), "Q2".to_string(), "Q3".to_string()];
        chart.series[0].values = vec![4.0, 8.0, 6.0];
        chart.series[1].values = vec![2.0, 6.0, 3.0];
        chart.category_axis_text_font_family = Some("Cambria".to_string());
        chart.value_axis_text_font_family = Some("Arial".to_string());
        let source = framed_chart_source(&chart, 480.0, 320.0);
        assert!(
            source.contains("font: (\"Cambria\""),
            "category font for {kind:?}"
        );
        assert!(
            source.contains("font: (\"Arial\""),
            "value font for {kind:?}"
        );
    }
}

#[test]
fn budget_scatter_marker_retains_its_declared_size() {
    let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
    let (doc, _) = crate::parser::Parser::parse(
        &crate::parser::xlsx::XlsxParser,
        data,
        &crate::config::ConvertOptions::default(),
    )
    .expect("the budget fixture parses");
    let source = generate_typst(&doc).unwrap().source;
    assert!(
        source.contains("circle(radius: 7pt,"),
        "the selected-month marker must retain its declared 14pt diameter"
    );
}

/// Native Excel's selected-month circle uses a worksheet raster grid. Odd
/// marker sizes move the ink center by half a sheet point in both directions;
/// changing the cached value would not reproduce this size-dependent phase.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn budget_compiled_marker_centers_match_native_placement_controls() {
    let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
    use crate::ir::ChartLine;
    let outline = |width| ChartLine::Explicit {
        width_pt: Some(width),
        color: None,
        alpha: None,
    };
    // Expected centers come from native exports that independently change
    // marker size, outline width/suppression, or both vertical anchor ends.
    for (size, line, shift_y, (native_x, native_y)) in [
        (13.0, outline(0.75), 0.0, (194.61, 310.83)),
        (14.0, outline(0.75), 0.0, (195.0, 311.22)),
        (15.0, outline(0.75), 0.0, (194.61, 310.83)),
        (20.0, outline(0.75), 0.0, (195.0, 311.22)),
        (14.0, ChartLine::Suppressed, 0.0, (195.78, 312.0)),
        (14.0, outline(0.0), 0.0, (195.0, 311.22)),
        (14.0, outline(1.49), 0.0, (195.0, 311.22)),
        (14.0, outline(1.5), 0.0, (195.78, 312.0)),
        (14.0, outline(2.0), 0.0, (195.78, 312.0)),
        (14.0, outline(2.5), 0.0, (195.0, 311.22)),
        (14.0, outline(3.5), 0.0, (195.78, 312.0)),
        (14.0, outline(0.75), 0.25, (195.0, 312.0)),
        (14.0, outline(0.75), 1.0, (195.0, 312.0)),
        (14.0, outline(0.75), 1.25, (195.0, 312.78)),
    ] {
        let (mut doc, _) = crate::parser::Parser::parse(
            &crate::parser::xlsx::XlsxParser,
            data,
            &crate::config::ConvertOptions::default(),
        )
        .expect("the public budget fixture parses");
        let mut changed: usize = 0;
        for page in &mut doc.pages {
            if let crate::ir::Page::Sheet(sheet) = page {
                for anchored in &mut sheet.charts {
                    for series in &mut anchored.chart.series {
                        if series.marker_symbol == Some(MarkerSymbol::Circle)
                            && series.marker_style.size_pt == Some(14.0)
                            && series.values.iter().any(|value| value.is_finite())
                        {
                            series.marker_style.size_pt = Some(size);
                            series.marker_style.line = line;
                            let placement = anchored.placement.as_mut().expect("anchored chart");
                            // Pagination already scales the anchor offset.
                            placement.y_offset_pt += shift_y * placement.print_scale;
                            changed += 1;
                        }
                    }
                }
            }
        }
        assert_eq!(changed, 1, "only the selected-month marker changes");
        let output = generate_typst(&doc).unwrap();
        let pages =
            crate::render::pdf::compiled_page_paint_sequences(&output.source, &output.images)
                .expect("the budget size control compiles");
        assert_eq!(pages.len(), 2);
        let circles: Vec<_> = pages[1]
            .iter()
            .filter(|paint| {
                let (left, top, right, bottom) = paint.bounds;
                paint.kind == crate::render::pdf::PaintedKind::Shape
                    && paint.rectangle_fill.is_none()
                    && left > 175.0
                    && right < 215.0
                    && top > 290.0
                    && bottom < 330.0
                    && (right - left - size * 0.78).abs() < 0.01
                    && (bottom - top - size * 0.78).abs() < 0.01
            })
            .collect();
        assert_eq!(circles.len(), 1, "one {size}pt plot circle");
        let (left, top, right, bottom) = circles[0].bounds;
        let actual = ((left + right) / 2.0, (top + bottom) / 2.0);
        assert!(
            (actual.0 - native_x).abs() < 0.01 && (actual.1 - native_y).abs() < 0.01,
            "{size}pt marker, {line:?}, anchor dy {shift_y}: native ({native_x}, {native_y}), actual {actual:?}"
        );
    }
}

/// Native gift-workbook size controls pin every plotted circle, including
/// zero values and different column bands, after source-font gutter recovery.
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn gift_compiled_marker_centers_match_native_size_controls() {
    let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1603_gift_budget.xlsx");
    for (size, native) in [
        (
            4.0,
            [
                (351.78000, 299.30000),
                (417.38000, 329.64000),
                (483.80000, 329.64000),
                (550.22000, 309.13997),
                (616.64000, 329.64000),
                (683.06000, 146.77999),
                (748.66000, 258.30001),
                (815.08000, 329.64000),
                (881.50000, 329.64000),
                (947.92000, 329.64000),
                (1014.34000, 329.64000),
                (1080.76000, 329.64000),
            ],
        ),
        (
            5.0,
            [
                (351.37000, 298.89003),
                (416.97000, 329.23003),
                (483.39000, 329.23003),
                (549.81000, 308.73000),
                (616.23000, 329.23003),
                (682.65000, 146.36999),
                (748.25000, 257.89000),
                (814.67000, 329.23003),
                (881.09000, 329.23003),
                (947.51000, 329.23003),
                (1013.93000, 329.23003),
                (1080.35000, 329.23003),
            ],
        ),
        (
            6.0,
            [
                (351.78000, 299.30003),
                (417.38000, 329.64003),
                (483.80000, 329.64003),
                (550.22000, 309.14000),
                (616.64000, 329.64003),
                (683.06000, 146.78000),
                (748.66000, 258.30002),
                (815.08000, 329.64003),
                (881.50000, 329.64003),
                (947.92000, 329.64003),
                (1014.34000, 329.64003),
                (1080.76000, 329.64003),
            ],
        ),
    ] {
        let (mut doc, _) = crate::parser::Parser::parse(
            &crate::parser::xlsx::XlsxParser,
            data,
            &crate::config::ConvertOptions::default(),
        )
        .expect("the public gift fixture parses");
        let mut changed: usize = 0;
        for page in &mut doc.pages {
            if let crate::ir::Page::Sheet(sheet) = page {
                for anchored in &mut sheet.charts {
                    for series in &mut anchored.chart.series {
                        if series.marker_symbol == Some(MarkerSymbol::Circle)
                            && series.marker_style.size_pt == Some(5.0)
                        {
                            series.marker_style.size_pt = Some(size);
                            changed += 1;
                        }
                    }
                }
            }
        }
        assert_eq!(changed, 1, "only the amount-spent marker changes");
        let output = generate_typst(&doc).unwrap();
        let pages =
            crate::render::pdf::compiled_page_paint_sequences(&output.source, &output.images)
                .expect("the gift marker control compiles");
        assert_eq!(pages.len(), 2);
        let mut circles: Vec<_> = pages[1]
            .iter()
            .filter(|paint| {
                let (left, top, right, bottom) = paint.bounds;
                paint.kind == crate::render::pdf::PaintedKind::Shape
                    && paint.rectangle_fill.is_none()
                    && left > 300.0
                    && right < 1125.0
                    && top > 120.0
                    && bottom < 340.0
                    && (right - left - size * 0.82).abs() < 0.01
                    && (bottom - top - size * 0.82).abs() < 0.01
            })
            .map(|paint| {
                let (left, top, right, bottom) = paint.bounds;
                ((left + right) / 2.0, (top + bottom) / 2.0)
            })
            .collect();
        circles.sort_by(|a, b| a.0.total_cmp(&b.0));
        assert_eq!(circles.len(), native.len());
        for (index, (actual, expected)) in circles.iter().zip(native).enumerate() {
            assert!(
                (actual.0 - expected.0).abs() < 0.015 && (actual.1 - expected.1).abs() < 0.015,
                "{size}pt marker {index}: native {expected:?}, actual {actual:?}"
            );
        }
    }
}

#[test]
fn budget_month_axis_preserves_its_single_trailing_space() {
    let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
    let (doc, _) = crate::parser::Parser::parse(
        &crate::parser::xlsx::XlsxParser,
        data,
        &crate::config::ConvertOptions::default(),
    )
    .expect("the budget fixture parses");
    let source = generate_typst(&doc).unwrap().source;
    assert!(
        source.contains(r#")[jan#" ";]"#),
        "single trailing category spaces must survive the centered Typst text boundary"
    );
    assert!(source.contains(r#")[year#"  ";]"#));
}

#[test]
fn category_axis_boundary_spaces_survive_each_axis_chart_renderer() {
    for kind in [
        ChartType::Bar,
        ChartType::Column,
        ChartType::Line,
        ChartType::Area,
        ChartType::Other(crate::ir::RADAR_CHART_LABEL.to_string()),
    ] {
        let mut chart = sized_bar_chart(11.0);
        chart.chart_type = kind.clone();
        chart.categories = [" Q1", "Q2 ", " Q3 ", "Q4", "Q5  ", " "]
            .into_iter()
            .map(str::to_string)
            .collect();
        for series in &mut chart.series {
            series.values = vec![4.0, 8.0, 6.0, 3.0, 7.0, 5.0];
        }
        let source = framed_chart_source(&chart, 480.0, 320.0);
        for label in [
            r#")[#" ";Q1]"#,
            r#")[Q2#" ";]"#,
            r#")[#" ";Q3#" ";]"#,
            ")[Q4]",
            r#")[Q5#"  ";]"#,
            r#")[#" ";]"#,
        ] {
            assert!(source.contains(label), "{kind:?} must preserve {label}");
        }
    }
}

/// Native Excel probes of the budget chart with only its category face/size
/// changed. Seats are in chart-local points before the worksheet's 0.78 scale.
/// Calibri's source metrics are available on every runner, including wasm.
#[test]
fn worksheet_line_category_baselines_follow_plot_space_and_font_size() {
    for kind in [ChartType::Line, ChartType::Area] {
        for (size, plot_height, native_baseline) in [
            (12.0, MONTHLY_BUDGET_LINE_PLOT_LAYOUT.height, 67.0),
            (15.0, MONTHLY_BUDGET_LINE_PLOT_LAYOUT.height, 70.0),
            (24.0, MONTHLY_BUDGET_LINE_PLOT_LAYOUT.height, 68.0),
            (12.0, 0.2, 40.0),
            (15.0, 0.2, 45.0),
            (24.0, 0.2, 59.0),
        ] {
            let mut chart = combo_line_and_scatter_chart();
            chart.chart_type = kind.clone();
            chart.series.truncate(1);
            chart.legend_position = LegendPosition::Right;
            chart.category_axis_text_font_family = Some("Calibri".to_string());
            chart.category_axis_text_style.size_pt = Some(size);
            chart.plot_area_layout = Some(crate::ir::ChartPlotAreaLayout {
                height: plot_height,
                ..MONTHLY_BUDGET_LINE_PLOT_LAYOUT
            });
            let source = framed_chart_source(&chart, 786.7143307086615, 75.54897637795276);
            let label = placed_box_holding(&source, "jan");
            assert!(
                (label.dy - native_baseline).abs() < 0.01,
                "{kind:?} {size}pt, plot height {plot_height}: native baseline {native_baseline}, got {label:?}"
            );
            let line = source.lines().find(|line| line.contains("[jan]")).unwrap();
            assert!(
                line.contains("top-edge: 0pt"),
                "dy must locate the baseline: {line}"
            );
        }
    }
}

#[test]
fn worksheet_line_category_cap_rounds_the_remaining_line_space() {
    let mut chart = combo_line_and_scatter_chart();
    chart.legend_position = LegendPosition::Right;
    chart.category_axis_text_font_family = Some("Calibri".to_string());
    chart.category_axis_text_style.size_pt = Some(15.0);
    chart.plot_area_layout = Some(crate::ir::ChartPlotAreaLayout {
        height: 0.8,
        ..MONTHLY_BUDGET_LINE_PLOT_LAYOUT
    });
    // One-factor native exports change only the drawing's bottom row offset.
    // Calibri crosses the next baseline at +0.4pt; Cambria does so at +0.6pt,
    // separating rounding the available line space from rounding the frame.
    for (increment, native_baseline) in [
        (0.0, 70.0),
        (0.2, 70.0),
        (0.4, 71.0),
        (0.6, 71.0),
        (1.0, 71.0),
        (1.4, 72.0),
    ] {
        let source = framed_chart_source(&chart, 786.7143307086615, 75.54897637795276 + increment);
        let label = placed_box_holding(&source, "jan");
        assert!(
            (label.dy - native_baseline).abs() < 0.01,
            "height +{increment}pt: expected native baseline {native_baseline}, got {label:?}"
        );
    }
}

#[test]
fn uncalibrated_line_category_labels_keep_the_existing_seat() {
    let mut chart = combo_line_and_scatter_chart();
    chart.legend_position = LegendPosition::Right;
    chart.category_axis_text_font_family = Some("Calibri".to_string());
    chart.category_axis_text_style.size_pt = Some(15.0);
    chart.plot_area_layout = Some(MONTHLY_BUDGET_LINE_PLOT_LAYOUT);
    for host in [
        crate::ir::ChartHost::WordProcessing,
        crate::ir::ChartHost::Presentation,
        crate::ir::ChartHost::SpreadsheetChartsheet,
    ] {
        chart.host = host;
        let source = framed_chart_source(&chart, 786.7143307086615, 75.54897637795276);
        let label = placed_box_holding(&source, "jan");
        assert!(
            (label.dy - 60.54897637795276).abs() < 0.01,
            "{host:?}: {label:?}"
        );
    }
    chart.host = crate::ir::ChartHost::Spreadsheet;
    let flowed = chart_source(chart.clone());
    assert!(
        !flowed
            .lines()
            .find(|line| line.contains("[jan]"))
            .unwrap()
            .contains("top-edge: 0pt")
    );
    chart.category_axis_text_style.size_pt = None;
    let implicit = framed_chart_source(&chart, 786.7143307086615, 75.54897637795276);
    assert!((placed_box_holding(&implicit, "jan").dy - 60.54897637795276).abs() < 0.01);
}

#[test]
fn category_line_spacing_quantizes_the_faces_leading_before_seating() {
    use crate::render::typst_gen::diagrams::line_category_baseline_pt;
    // Native Arial exports distinguish its nonzero hhea leading from the
    // zero-leading Calibri controls. Metrics are the source Office face's;
    // supplying them directly keeps this numeric layout check font-portable.
    let metrics = (1854.0 / 2048.0, 434.0 / 2048.0, 67.0 / 2048.0);
    let height = 75.54897637795276;
    for (size, fraction, native) in [
        (12.0, 0.559_675_203_526_960_8, 67.0),
        (15.0, 0.559_675_203_526_960_8, 71.0),
        (18.0, 0.559_675_203_526_960_8, 70.0),
        (24.0, 0.559_675_203_526_960_8, 70.0),
        (12.0, 0.2, 39.0),
        (15.0, 0.2, 44.0),
        (24.0, 0.2, 58.0),
    ] {
        let baseline = line_category_baseline_pt(
            height,
            (0.07430787730224597 + fraction) * height,
            size,
            metrics,
        );
        assert!(
            (baseline - native).abs() < 0.01,
            "Arial {size}pt, plot height {fraction}: native {native}, got {baseline}"
        );
    }
}

#[test]
fn bar_category_label_baseline_matches_measured_native_interior_rows() {
    use crate::render::typst_gen::diagrams::bar_category_label_baseline_pt;
    // Native Excel baselines for `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`'s
    // two worksheet bar charts, regular Trebuchet MS in Excel 16.112 (#1621).
    // Baselines are unscaled sheet points. Three distinct
    // plot geometries at 10pt (income @ 5 and 6 categories, expense @ 6) and one
    // chart at 8/14pt all land on `floor(sheet_frame_top_pt + row_top + row / 2)
    // + K`, with `K` depending only on the declared size — see
    // `assets/validation/issue-1621/README.md` and `measurements.json`.
    const SHEET_FRAME_TOP_PT: f64 = 159.5;
    for (row_top, row, size_pt, native_sheet) in [
        // Income, 5 categories, plot_y=71.85567962392459, row=23.122867122695773.
        (94.97854674662037, 23.122867122695773, 8.0, 270.0), // from savings
        (118.10141386931613, 23.122867122695773, 8.0, 293.0), // family help
        (94.97854674662037, 23.122867122695773, 10.0, 270.0), // from savings
        (118.10141386931613, 23.122867122695773, 10.0, 293.0), // family help
        (141.22428099201193, 23.122867122695773, 10.0, 316.0), // wages (after-tax)
        (94.97854674662037, 23.122867122695773, 14.0, 271.0), // from savings
        (118.10141386931613, 23.122867122695773, 14.0, 294.0), // family help
        // Expense, 6 categories, plot_y=70.87126605261206, row=19.59719379268398.
        (90.46845984529602, 19.59719379268398, 10.0, 263.0), // discretionary
        (110.06565363798, 19.59719379268398, 10.0, 283.0),   // transportation
        (129.662847430664, 19.59719379268398, 10.0, 302.0),  // books & supplies
        (149.26004122334797, 19.59719379268398, 10.0, 322.0), // tuition & fees
        // Income widened to 6 categories (category-count probe), row=19.26905593557981.
        (91.1247355595044, 19.26905593557981, 10.0, 264.0), // other
        (110.3937914950842, 19.26905593557981, 10.0, 283.0), // from savings
        (129.662847430664, 19.26905593557981, 10.0, 302.0), // family help
        (148.93190330624383, 19.26905593557981, 10.0, 322.0), // wages (after-tax)
    ] {
        let local_dy = bar_category_label_baseline_pt(SHEET_FRAME_TOP_PT, row_top, row, size_pt)
            .unwrap_or_else(|| panic!("{size_pt}pt is a measured size"));
        let got_sheet = local_dy + SHEET_FRAME_TOP_PT;
        assert!(
            (got_sheet - native_sheet).abs() < 0.001,
            "row_top={row_top}, row={row}, size={size_pt}: expected sheet {native_sheet}, got {got_sheet}"
        );
    }
}

#[test]
fn bar_category_label_baseline_keeps_no_rule_for_an_unmeasured_size() {
    use crate::render::typst_gen::diagrams::bar_category_label_baseline_pt;
    assert_eq!(
        bar_category_label_baseline_pt(159.5, 110.0, 20.0, 11.0),
        None
    );
}

#[test]
fn a_worksheet_bar_chart_seats_only_interior_category_labels_on_the_quantized_baseline() {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.series.truncate(1);
    // A lone series would otherwise stand in as an automatic chart title,
    // shifting the plot rectangle down by a title band this test's
    // independent geometry check doesn't otherwise account for.
    chart.auto_title_deleted = true;
    chart.categories = vec![
        "c0".into(),
        "c1".into(),
        "c2".into(),
        "c3".into(),
        "c4".into(),
    ];
    chart.series[0].values = vec![1.0, 2.0, 3.0, 4.0, 5.0];
    chart.host = crate::ir::ChartHost::Spreadsheet;
    chart.category_axis_text_style.size_pt = Some(10.0);
    // Axis overrides resolve to the measured face despite different chart defaults.
    chart.text_font_family = Some("Arial".into());
    chart.text_style.bold = Some(true);
    chart.category_axis_text_font_family = Some("Trebuchet MS".into());
    chart.category_axis_text_style.bold = Some(false);

    let frame = (400.0, 300.0);
    const SHEET_FRAME_TOP_PT: f64 = 20.5;
    let mut source = String::new();
    generate_sheet_chart_in(&mut source, &chart, frame, (50.0, SHEET_FRAME_TOP_PT));

    let (_, top, _, bottom) = axis_plot_rect(&chart, frame, false);
    let plot_h: f64 = bottom - top;
    let row: f64 = plot_h / chart.categories.len() as f64;

    for (cat_index, label) in ["c0", "c1", "c2", "c3", "c4"].into_iter().enumerate() {
        let line: &str = source
            .lines()
            .find(|line| line.contains(&format!("[{label}]")))
            .unwrap_or_else(|| panic!("nothing prints {label} in:\n{source}"));
        let placed = placed_box_holding(&source, label);
        // Preserve #1620's native 10pt horizontal clearance in both baseline paths.
        assert!(
            (placed.dx - 0.2610 / 0.78).abs() < 0.01,
            "{label}: horizontal clearance must survive the baseline correction: {placed:?}"
        );
        let is_edge_row: bool = cat_index == 0 || cat_index + 1 == chart.categories.len();
        if is_edge_row {
            assert!(
                !line.contains("top-edge: 0pt, bottom-edge: 0pt"),
                "edge row {label} keeps its existing centred seat: {line}"
            );
            continue;
        }
        assert!(
            line.contains("top-edge: 0pt, bottom-edge: 0pt"),
            "interior row {label} seats directly on its quantized baseline: {line}"
        );
        let row_top: f64 = plot_h - (cat_index as f64 + 1.0) * row;
        let expected = crate::render::typst_gen::diagrams::bar_category_label_baseline_pt(
            SHEET_FRAME_TOP_PT,
            row_top,
            row,
            10.0,
        )
        .expect("10pt is a measured size");
        assert!(
            (placed.dy - expected).abs() < 0.01,
            "{label}: expected {expected}, got {placed:?}"
        );
    }
}

#[test]
fn worksheet_bar_category_baselines_keep_unmeasured_faces_and_weights_centred() {
    for (family, inherited_bold, axis_bold) in [
        ("Arial", false, None),
        ("Calibri", false, None),
        ("Trebuchet MS", false, Some(true)),
        ("Trebuchet MS", true, None),
    ] {
        let mut chart = two_series_bar_chart(Vec::new());
        chart.series.truncate(1);
        chart.categories = vec!["Rent".into(), "Food".into(), "Travel".into()];
        chart.series[0].values = vec![900.0, 300.0, 120.0];
        chart.host = crate::ir::ChartHost::Spreadsheet;
        chart.text_font_family = Some(family.into());
        chart.text_style.bold = Some(inherited_bold);
        chart.category_axis_text_style.bold = axis_bold;
        chart.category_axis_text_style.size_pt = Some(10.0);

        let mut source = String::new();
        generate_sheet_chart_in(&mut source, &chart, (400.0, 300.0), (50.0, 20.5));
        let line: &str = source.lines().find(|line| line.contains("[Food]")).unwrap();
        assert!(
            !line.contains("top-edge: 0pt, bottom-edge: 0pt"),
            "unmeasured {family}, inherited bold {inherited_bold}, axis bold {axis_bold:?}: {line}"
        );
    }
}

#[test]
fn a_worksheet_bar_chart_at_an_unmeasured_size_keeps_every_labels_centred_seat() {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.series.truncate(1);
    chart.categories = vec!["c0".into(), "c1".into(), "c2".into()];
    chart.series[0].values = vec![1.0, 2.0, 3.0];
    chart.host = crate::ir::ChartHost::Spreadsheet;
    chart.category_axis_text_style.size_pt = Some(11.0); // never measured against native (#1621)

    let mut source = String::new();
    generate_sheet_chart_in(&mut source, &chart, (400.0, 300.0), (50.0, 20.5));
    let line = source
        .lines()
        .find(|line| line.contains("[c1]"))
        .expect("c1's label is emitted");
    assert!(
        !line.contains("top-edge: 0pt, bottom-edge: 0pt"),
        "interior row at an unmeasured size keeps its existing seat: {line}"
    );
}

#[test]
fn declared_marker_paint_reaches_plot_and_legend_across_chart_families() {
    for kind in [
        ChartType::Line,
        ChartType::Area,
        ChartType::Column,
        ChartType::Other(crate::ir::RADAR_CHART_LABEL.to_string()),
    ] {
        for size in [8.0, 14.0, 28.0] {
            let mut chart = two_series_bar_chart(Vec::new());
            chart.chart_type = kind.clone();
            chart.categories = vec!["Q1".into(), "Q2".into(), "Q3".into()];
            chart.series.truncate(1);
            let series = &mut chart.series[0];
            series.values = vec![4.0, 8.0, 6.0];
            series.plot_type = if kind == ChartType::Column {
                Some(ChartType::Line)
            } else {
                None
            };
            series.marker_symbol = Some(crate::ir::MarkerSymbol::Circle);
            series.marker_style = crate::ir::ChartMarkerStyle {
                size_pt: Some(size),
                fill_mode: crate::ir::ChartFillMode::Explicit,
                fill: Some(crate::ir::Color::new(0, 136, 137)),
                fill_alpha: None,
                line: crate::ir::ChartLine::Explicit {
                    width_pt: Some(0.75),
                    color: Some(crate::ir::Color::new(17, 34, 51)),
                    alpha: None,
                },
            };
            let source = chart_source(chart);
            let paint: &str = "fill: rgb(0, 136, 137), stroke: 0.75pt + rgb(17, 34, 51))";
            let plotted: String = format!("circle(radius: {}pt, {paint}", format_f64(size / 2.0));
            assert_eq!(
                source.matches(&plotted).count(),
                3,
                "three plotted markers at the declared size for {kind:?}: {source}"
            );
            // The legend key carries the same paint at the legend's own marker
            // allowance (#1617); the chart declares no legend size here. A key
            // under the allowance is the plotted markup itself, so it counts
            // among the three matches above.
            let key_size: f64 = size.min(legend_marker_cap_pt(CHART_DEFAULT_TEXT_PT));
            let key: String = format!("circle(radius: {}pt, {paint}", format_f64(key_size / 2.0));
            let expected_keys: usize = if key_size < size { 1 } else { 4 };
            assert_eq!(
                source.matches(&key).count(),
                expected_keys,
                "one legend key marker carrying the declared paint for {kind:?}: {source}"
            );
        }
    }
}

#[test]
fn marker_fill_and_outline_suppression_reach_the_plot_independently() {
    let mut chart = two_series_bar_chart(Vec::new());
    chart.categories = vec!["Q1".into(), "Q2".into()];
    chart.series[0].values = vec![4.0, 8.0];
    chart.chart_type = ChartType::Line;
    chart.series.truncate(1);
    chart.series[0].marker_symbol = Some(crate::ir::MarkerSymbol::Circle);
    chart.series[0].marker_style.fill_mode = crate::ir::ChartFillMode::Suppressed;
    chart.series[0].marker_style.line = crate::ir::ChartLine::Explicit {
        width_pt: Some(2.0),
        color: Some(crate::ir::Color::new(17, 34, 51)),
        alpha: None,
    };
    let outlined = chart_source(chart.clone());
    assert!(outlined.contains("fill: none, stroke: 2pt + rgb(17, 34, 51)"));
    chart.series[0].marker_style.fill_mode = crate::ir::ChartFillMode::Explicit;
    chart.series[0].marker_style.fill = Some(crate::ir::Color::new(0, 136, 137));
    chart.series[0].marker_style.line = crate::ir::ChartLine::Suppressed;
    let filled = chart_source(chart);
    assert!(filled.contains("fill: rgb(0, 136, 137), stroke: none"));
}

/// A selected-period point can overlap the zero/category rule. Native Excel
/// paints that rule first; an opaque marker covers it and a hollow one lets
/// it show through. Inspect compiled paint order, not markup order (#1578).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn category_axis_paints_before_overlapping_line_plot_markers() {
    use crate::ir::{ChartFillMode, ChartLine};
    use crate::render::pdf::{PaintedKind, compiled_paint_sequence};

    let mut failures = Vec::new();
    for value in [0.0, 0.5] {
        for (symbol, fill_mode) in [
            (MarkerSymbol::Circle, ChartFillMode::Explicit),
            (MarkerSymbol::Circle, ChartFillMode::Suppressed),
            (MarkerSymbol::Off, ChartFillMode::Explicit),
        ] {
            let mut chart = combo_line_and_scatter_chart();
            chart.series[0].values = vec![0.0, 30.0, -10.0];
            chart.series[1].values = vec![value];
            chart.series[1].marker_symbol = Some(symbol);
            chart.series[1].marker_style = crate::ir::ChartMarkerStyle {
                size_pt: Some(14.0),
                fill_mode,
                fill: Some(crate::ir::Color::new(204, 34, 68)),
                line: ChartLine::Explicit {
                    width_pt: Some(0.75),
                    color: Some(crate::ir::Color::new(0, 136, 137)),
                    alpha: Some(1.0),
                },
                ..Default::default()
            };
            chart.value_axis_min = Some(-20.0);
            chart.value_axis_max = Some(40.0);
            chart.major_gridline_line = ChartLine::Suppressed;
            chart.category_axis_line = ChartLine::Explicit {
                width_pt: Some(0.25),
                color: Some(crate::ir::Color::new(191, 191, 191)),
                alpha: Some(0.5),
            };
            let source = chart_source(chart);
            let paints = compiled_paint_sequence(&source, &[], 0).unwrap();
            let axes: Vec<_> = paints
                .iter()
                .enumerate()
                .filter(|(_, item)| {
                    let (x0, y0, x1, y1) = item.bounds;
                    item.kind == PaintedKind::Shape && x1 - x0 > 100.0 && (y1 - y0).abs() < 0.01
                })
                .collect();
            assert_eq!(axes.len(), 1, "the declared category rule remains present");
            let markers: Vec<_> = paints
                .iter()
                .enumerate()
                .filter(|(_, item)| {
                    let (x0, y0, x1, y1) = item.bounds;
                    item.kind == PaintedKind::Shape
                        && (x1 - x0 - 14.0).abs() < 0.01
                        && (y1 - y0 - 14.0).abs() < 0.01
                })
                .collect();
            if symbol == MarkerSymbol::Off {
                assert!(markers.is_empty(), "disabled markers must remain absent");
                continue;
            }
            assert_eq!(markers.len(), 1, "the declared 14pt marker is retained");
            let (axis_index, axis) = axes[0];
            let (marker_index, marker) = markers[0];
            assert!(axis.bounds.1 > marker.bounds.1 && axis.bounds.1 < marker.bounds.3);
            if axis_index >= marker_index {
                failures.push(format!(
                    "value={value}, fill={fill_mode:?}: axis paint {axis_index} follows marker {marker_index}"
                ));
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

/// A line laid over columns can run along the category axis: the #982 workbook
/// spends nothing in most months, so its `Amount Spent` line lies on the zero
/// rule there. Native Excel strokes that rule after the columns and before the
/// line, so the line's own colour shows where the two coincide. Inspect the
/// compiled paint order, not markup order (#1604).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn category_axis_paints_between_columns_and_overlaid_line_series() {
    use crate::ir::ChartLine;
    use crate::render::pdf::{PaintedKind, PaintedPrimitive, compiled_paint_sequence};

    const AXIS_WIDTH_PT: f64 = 0.6;
    const LINE_WIDTH_PT: f64 = 1.84;
    let stroke_width = |item: &PaintedPrimitive| -> f64 {
        item.stroke
            .as_ref()
            .map(|stroke| stroke.thickness_pt)
            .unwrap_or(0.0)
    };
    let mut failures = Vec::new();
    for chart_type in [ChartType::Column, ChartType::Bar] {
        for line_values in [vec![0.0, 0.0], vec![0.0, 75.0]] {
            let mut chart = combo_budget_chart();
            chart.chart_type = chart_type.clone();
            chart.has_legend = false;
            chart.series[2].values = line_values.clone();
            chart.series[2].line_width_pt = Some(LINE_WIDTH_PT);
            chart.category_axis_line = ChartLine::Explicit {
                width_pt: Some(AXIS_WIDTH_PT),
                color: Some(crate::ir::Color::new(217, 217, 217)),
                alpha: None,
            };
            chart.value_axis_line = ChartLine::Suppressed;
            chart.major_gridline_line = ChartLine::Suppressed;
            let source = chart_source(chart);
            let paints = compiled_paint_sequence(&source, &[], 0).unwrap();
            let case = format!("{chart_type:?} with line values {line_values:?}");

            let axes: Vec<(usize, &PaintedPrimitive)> = paints
                .iter()
                .enumerate()
                .filter(|(_, item)| {
                    item.kind == PaintedKind::Shape
                        && (stroke_width(item) - AXIS_WIDTH_PT).abs() < 0.01
                })
                .collect();
            assert_eq!(
                axes.len(),
                1,
                "{case}: the declared category rule is drawn once"
            );
            let lines: Vec<(usize, &PaintedPrimitive)> = paints
                .iter()
                .enumerate()
                .filter(|(_, item)| {
                    item.kind == PaintedKind::Shape
                        && (stroke_width(item) - LINE_WIDTH_PT).abs() < 0.01
                })
                .collect();
            assert_eq!(
                lines.len(),
                1,
                "{case}: the overlaid series is stroked once"
            );
            let last_column: usize = paints
                .iter()
                .enumerate()
                .filter(|(_, item)| {
                    item.kind == PaintedKind::Shape
                        && item.rectangle_fill.is_some()
                        && item.bounds.2 - item.bounds.0 > 1.0
                        && item.bounds.3 - item.bounds.1 > 1.0
                })
                .map(|(index, _)| index)
                .max()
                .expect("the stacked budget columns are painted");

            let (axis_index, axis) = axes[0];
            let (line_index, line) = lines[0];
            // The zero-valued points put the series on the rule itself.
            let reach: f64 = LINE_WIDTH_PT / 2.0;
            let overlaps: bool = line.bounds.0 - reach <= axis.bounds.2
                && line.bounds.2 + reach >= axis.bounds.0
                && line.bounds.1 - reach <= axis.bounds.3
                && line.bounds.3 + reach >= axis.bounds.1;
            assert!(overlaps, "{case}: the series lies on the category rule");
            if last_column >= axis_index {
                failures.push(format!(
                    "{case}: column paint {last_column} follows the axis paint {axis_index}"
                ));
            }
            if axis_index >= line_index {
                failures.push(format!(
                    "{case}: axis paint {axis_index} follows the series paint {line_index}"
                ));
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

/// Read actual compiled plot and legend strokes from declared series XML.
/// Native Excel's cap/join controls also establish the omitted miter limit (#1590).
#[cfg(not(target_arch = "wasm32"))]
#[test]
fn declared_series_stroke_geometry_reaches_plots_and_legends() {
    use crate::parser::drawingml::SchemeColors;
    use crate::render::pdf::compiled_page_paint_sequences;
    use typst::visualize::{LineCap, LineJoin};

    let colors = std::collections::HashMap::new();
    let aliases = std::collections::HashMap::new();
    let scheme = SchemeColors {
        colors: &colors,
        aliases: &aliases,
    };
    let cases = [
        ("rnd", "<a:round/>", LineCap::Round, LineJoin::Round, None),
        ("flat", "<a:round/>", LineCap::Butt, LineJoin::Round, None),
        ("sq", "<a:bevel/>", LineCap::Square, LineJoin::Bevel, None),
        (
            "rnd",
            "<a:miter lim=\"800000\"/>",
            LineCap::Round,
            LineJoin::Miter,
            Some(8.0),
        ),
        (
            "flat",
            "<a:miter lim=\"200000\"/>",
            LineCap::Butt,
            LineJoin::Miter,
            Some(2.0),
        ),
        (
            "sq",
            "<a:miter/>",
            LineCap::Square,
            LineJoin::Miter,
            Some(8.0),
        ),
    ];
    let mut pages = Vec::new();
    let mut expected = Vec::new();
    for (index, (cap, join, expected_cap, expected_join, limit)) in cases.iter().enumerate() {
        let width = if index % 2 == 0 { 2.25 } else { 3.5 };
        let emu = width * 12700.0;
        let xml = format!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart><c:ser>
          <c:tx><c:v>Series</c:v></c:tx>
          <c:spPr><a:ln w="{emu}" cap="{cap}"><a:solidFill><a:srgbClr val="CC2244"/></a:solidFill>{join}</a:ln></c:spPr>
          <c:marker><c:symbol val="none"/><c:spPr><a:ln w="12700" cap="flat"><a:bevel/></a:ln></c:spPr></c:marker>
          <c:dPt><c:idx val="1"/><c:spPr><a:ln w="12700" cap="flat"><a:bevel/></a:ln></c:spPr></c:dPt>
          <c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
          <c:val><c:numLit><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt><c:pt idx="2"><c:v>15</c:v></c:pt></c:numLit></c:val>
          </c:ser></c:lineChart></c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>"#
        );
        let parsed = crate::parser::chart::parse_chart_xml(&xml, &scheme).unwrap();
        for family in [
            ChartType::Line,
            ChartType::Other(crate::ir::RADAR_CHART_LABEL.to_string()),
            ChartType::Column,
        ] {
            let mut chart = parsed.clone();
            chart.chart_type = family.clone();
            if family == ChartType::Column {
                chart.series[0].plot_type = Some(ChartType::Line);
                let mut columns = parsed.series[0].clone();
                columns.name = Some("Columns".to_string());
                columns.values = vec![12.0, 18.0, 9.0];
                columns.line_width_pt = None;
                columns.plot_type = None;
                chart.series.insert(0, columns);
            }
            pages.push(make_flow_page(vec![Block::Chart(Box::new(chart))]));
            expected.push((
                family,
                *cap,
                *join,
                *expected_cap,
                *expected_join,
                *limit,
                width,
            ));
        }
    }
    let source = generate_typst(&make_doc(pages)).unwrap().source;
    let painted_pages = compiled_page_paint_sequences(&source, &[]).unwrap();
    assert_eq!(painted_pages.len(), expected.len());
    let mut failures = Vec::new();
    for (paints, (family, cap, join, expected_cap, expected_join, limit, width)) in
        painted_pages.iter().zip(expected)
    {
        let strokes: Vec<_> = paints
            .iter()
            .filter_map(|paint| paint.stroke.clone())
            .filter(|stroke| (stroke.thickness_pt - width).abs() < 1e-9)
            .collect();
        assert_eq!(
            strokes.len(),
            2,
            "{family:?}: plotted path and matching legend sample retain their width"
        );
        for stroke in strokes {
            if stroke.cap != expected_cap
                || stroke.join != expected_join
                || limit.is_some_and(|expected| (stroke.miter_limit - expected).abs() > 1e-9)
            {
                failures.push(format!("{family:?} cap={cap} join={join}: {stroke:?}"));
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

/// The `Success Ratios` line chart of the #1220 deck, reduced to its two axis
/// groups: a profit margin read against a 0..20% primary axis and an acid-test
/// ratio read against a 0..7 secondary axis on the right (issue #1374).
fn dual_axis_line_chart() -> Chart {
    let mut chart: Chart = two_series_bar_chart(Vec::new());
    chart.chart_type = ChartType::Line;
    chart.has_legend = false;
    chart.categories = vec![
        "Year 1".to_string(),
        "Year 2".to_string(),
        "Year 3".to_string(),
    ];
    chart.series[0].name = Some("Profit Margin".to_string());
    chart.series[0].values = vec![0.12, 0.1495, 0.1766];
    chart.series[0].number_format = Some("0.00%".to_string());
    chart.series[1].name = Some("ACID Test".to_string());
    chart.series[1].values = vec![2.34, 3.66, 6.77];
    chart.series[1].value_axis = crate::ir::ChartValueAxisRole::Secondary;
    chart.value_axis_max = Some(0.2);
    chart.value_axis_number_format = Some("0%".to_string());
    chart.value_axis_major_tick_mark = AxisTickMark::None;
    chart.secondary_value_axis = Some(crate::ir::ChartSecondaryValueAxis {
        side: crate::ir::ValueAxisSide::Right,
        deleted: false,
        number_format: None,
        major_unit: None,
        min: None,
        max: Some(7.0),
        major_tick_mark: AxisTickMark::None,
        line: crate::ir::ChartLine::Suppressed,
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
    });
    chart
}

/// Every value tick label the source prints at the default text size, with
/// the x offset of the box holding it. A value label is centred on its
/// gridline in a box of stated height; a category label's box states none,
/// which is what tells the two apart.
fn emitted_value_labels(source: &str) -> Vec<(f64, String)> {
    let marker: String = format!("text(size: {}pt)[", format_f64(CHART_DEFAULT_TEXT_PT));
    source
        .lines()
        .filter(|line| {
            line.contains("#place") && line.contains("box(width: ") && line.contains("height: ")
        })
        .filter_map(|line| {
            let label: &str = line.rsplit_once(marker.as_str())?.1.split_once(']')?.0;
            Some((
                leading_pt(line.split_once("dx: ")?.1)?,
                label.replace('\\', ""),
            ))
        })
        .collect()
}

/// The points of the `index`-th emitted `path(...)`, in placement coordinates.
fn emitted_path_points_of(source: &str, index: usize) -> Vec<(f64, f64)> {
    let line: &str = source
        .lines()
        .filter(|line| line.contains("curve(stroke:"))
        .nth(index)
        .unwrap_or_default();
    line.match_indices("(")
        .filter_map(|(start, _)| {
            let body: &str = line[start + 1..].split_once(')')?.0;
            let (x, y) = body.split_once(", ")?;
            Some((
                x.strip_suffix("pt")?.parse().ok()?,
                y.strip_suffix("pt")?.parse().ok()?,
            ))
        })
        .collect()
}

/// Each series is read against its own axis: the acid test's 6.77 sits at
/// 6.77/7 of the plot height, and the profit margin's 0.1766 at 0.1766/0.2 of
/// it. Before this both were read against one collapsed 0..7 scale, which
/// drew the margin as a flat line along the plot floor (issue #1374).
#[test]
fn a_secondary_axis_series_is_scaled_against_its_own_axis() {
    let source: String = chart_source(dual_axis_line_chart());
    let (_, plot_y, _, plot_h) = plot_rect(&emitted_lines(&source));

    let margin: Vec<(f64, f64)> = emitted_path_points_of(&source, 0);
    let acid: Vec<(f64, f64)> = emitted_path_points_of(&source, 1);
    assert_eq!(margin.len(), 3, "three margin points in:\n{source}");
    assert_eq!(acid.len(), 3, "three acid-test points in:\n{source}");

    let expected_margin_y: f64 = plot_y + (1.0 - 0.1766 / 0.2) * plot_h;
    let expected_acid_y: f64 = plot_y + (1.0 - 6.77 / 7.0) * plot_h;
    assert!(
        (margin[2].1 - expected_margin_y).abs() < 0.05,
        "the margin's 17.66% sits at {expected_margin_y:.2}pt on the 0..20% axis, got {:.2}pt",
        margin[2].1
    );
    assert!(
        (acid[2].1 - expected_acid_y).abs() < 0.05,
        "the acid test's 6.77 sits at {expected_acid_y:.2}pt on the 0..7 axis, got {:.2}pt",
        acid[2].1
    );
}

/// The primary axis prints its own 0%..20% ticks on the left and the
/// secondary its 0..7 on the right of the plot, each in its own format.
#[test]
fn each_value_axis_prints_its_own_ticks_on_its_own_side() {
    let source: String = chart_source(dual_axis_line_chart());
    let (plot_x, _, plot_w, _) = plot_rect(&emitted_lines(&source));
    let labels: Vec<(f64, String)> = emitted_value_labels(&source);

    let left: Vec<&str> = labels
        .iter()
        .filter(|(dx, _)| *dx < plot_x)
        .map(|(_, label)| label.as_str())
        .collect();
    let right: Vec<&str> = labels
        .iter()
        .filter(|(dx, _)| *dx >= plot_x + plot_w)
        .map(|(_, label)| label.as_str())
        .collect();
    assert_eq!(
        left,
        ["0%", "5%", "10%", "15%", "20%"],
        "left labels in:\n{source}"
    );
    assert_eq!(
        right,
        ["0", "1", "2", "3", "4", "5", "6", "7"],
        "right labels in:\n{source}"
    );
}

/// A framed chart keeps its secondary labels inside the frame: the plot gives
/// up a right-hand gutter for them, as it always has a left-hand one.
#[test]
fn a_framed_chart_reserves_a_gutter_for_its_secondary_axis_labels() {
    let (frame_w, frame_h) = (400.0, 200.0);
    let source: String = framed_chart_source(&dual_axis_line_chart(), frame_w, frame_h);
    let (plot_x, _, plot_w, _) = plot_rect(&emitted_lines(&source));

    let right_boxes: Vec<PlacedBox> = ["0", "7"]
        .iter()
        .map(|label| placed_box_holding(&source, label))
        .collect();
    for placed in right_boxes {
        assert!(
            placed.dx >= plot_x + plot_w,
            "a secondary label starts right of the plot edge {:.2}pt, got {:.2}pt in:\n{source}",
            plot_x + plot_w,
            placed.dx
        );
        assert!(
            placed.dx + placed.width <= frame_w + 1e-6,
            "a secondary label box ends inside the {frame_w}pt frame, got {:.2}pt in:\n{source}",
            placed.dx + placed.width
        );
    }
}

/// A chart that declares no `<c:legend>` reserves no legend band, so its plot
/// runs to the frame's right gutter.
///
/// The line renderer built its `LegendBox` from `<c:legendPos>` alone while
/// `has_legend` only gated whether entries were drawn, so a legendless line
/// chart still gave up the column a right legend would take and every
/// right-hand element sat that far left — 94pt of it on page 11 of the #1220
/// deck (issue #1649). The axis, radar and pie renderers already reclaimed it
/// (issue #762).
#[test]
fn a_line_chart_without_a_legend_reserves_no_legend_band() {
    // Both families the line renderer draws, at two frame widths: the band a
    // legend would take is a fixed column, so a renderer that merely scaled
    // the plot down would pass one width and fail the other.
    for chart_type in [ChartType::Line, ChartType::Area] {
        for frame_w in [400.0_f64, 560.0_f64] {
            let frame_h: f64 = 200.0;

            let mut legendless: Chart = dual_axis_line_chart();
            legendless.chart_type = chart_type.clone();
            legendless.legend_position = LegendPosition::Right;
            assert!(
                !legendless.has_legend,
                "the #1220 `Success Ratios` fixture declares no legend"
            );
            let without: String = framed_chart_source(&legendless, frame_w, frame_h);
            let (plot_x, _, plot_w, _) = plot_rect(&emitted_lines(&without));

            // The only thing between the plot and the frame edge is the
            // secondary axis' own label gutter, which mirrors the primary's on
            // the left. Anything wider is a band reserved for a legend that
            // does not exist.
            let right_gutter: f64 = frame_w - (plot_x + plot_w);
            assert!(
                (right_gutter - plot_x).abs() < 1e-6,
                "{chart_type:?} at {frame_w}pt: the right gutter is the secondary axis' \
                 {plot_x:.2}pt label band, got {right_gutter:.2}pt in:\n{without}"
            );

            // Triangulation: a declared legend still takes its column, and it
            // takes the same column at either frame width.
            let mut legended: Chart = dual_axis_line_chart();
            legended.chart_type = chart_type.clone();
            legended.legend_position = LegendPosition::Right;
            legended.has_legend = true;
            let with: String = framed_chart_source(&legended, frame_w, frame_h);
            let (_, _, legended_plot_w, _) = plot_rect(&emitted_lines(&with));
            assert!(
                plot_w > legended_plot_w,
                "{chart_type:?} at {frame_w}pt: a declared right legend still reserves a \
                 column: {plot_w:.2}pt vs {legended_plot_w:.2}pt"
            );
        }
    }
}

/// The reclaimed band is the legend's own column, not a share of the frame:
/// it is the same width whatever the frame measures.
#[test]
fn the_reclaimed_line_legend_band_is_a_fixed_column() {
    let bands: Vec<f64> = [400.0_f64, 560.0_f64]
        .into_iter()
        .map(|frame_w| {
            let mut legended: Chart = dual_axis_line_chart();
            legended.legend_position = LegendPosition::Right;
            legended.has_legend = true;
            let with: String = framed_chart_source(&legended, frame_w, 200.0);
            let without: String = framed_chart_source(&dual_axis_line_chart(), frame_w, 200.0);
            plot_rect(&emitted_lines(&without)).2 - plot_rect(&emitted_lines(&with)).2
        })
        .collect();

    assert!(
        (bands[0] - bands[1]).abs() < 1e-6,
        "the legend column is frame-independent, got {:.2}pt and {:.2}pt",
        bands[0],
        bands[1]
    );
    assert!(
        bands[0] > 0.0,
        "a declared right legend reserves a column, got {:.2}pt",
        bands[0]
    );
}

/// A line chart's category label is laid out in its own category band, so a
/// label wider than the retired fixed 24pt box stays on the one line every
/// reference exporter prints it on.
///
/// On page 11 of the `GENERAL SERVICES.pptx` deck from #1220 the three
/// `Year N` labels are set at 11.97pt over bands more than four times the old
/// box wide, and both the LibreOffice reference and a native PowerPoint export
/// print each on one line; the fixed box broke every one into `Year` over its
/// digit. The column renderer has always laid its labels out in `row`, the
/// band width, so this is the same law on the other renderer (issue #1650).
#[test]
fn a_line_chart_lays_each_category_label_out_in_its_band() {
    // Two frame widths and two category counts: the band is the plot divided
    // by the categories, so a renderer that merely widened the constant, or
    // that scaled it with the frame alone, passes at most one of these.
    for (frame_w, categories) in [(400.0_f64, 3usize), (560.0_f64, 4usize)] {
        let mut chart: Chart = dual_axis_line_chart();
        chart.categories = (1..=categories)
            .map(|year| format!("Year {year}"))
            .collect();
        for series in &mut chart.series {
            series.values = (0..categories).map(|index| index as f64).collect();
        }
        chart.category_axis_text_style.size_pt = Some(12.0);

        let source: String = framed_chart_source(&chart, frame_w, 200.0);
        let (plot_x, _, plot_w, _) = plot_rect(&emitted_lines(&source));
        let band_w: f64 = plot_w / categories as f64;
        assert!(
            band_w > 24.0,
            "the fixture must give each label a band wider than the retired box \
             for this to distinguish them, got {band_w:.2}pt"
        );

        for (index, category) in chart.categories.iter().enumerate() {
            let placed: PlacedBox = placed_box_holding(&source, category);
            assert!(
                same_length(placed.width, band_w),
                "{categories} categories in a {frame_w}pt frame: `{category}` must be \
                 laid out in its {band_w:.2}pt band, got a {:.2}pt box in:\n{source}",
                placed.width
            );
            // The band starts where the previous one ends, so the label stays
            // centred on its point while the box grows around it.
            let band_start: f64 = plot_x + index as f64 * band_w;
            assert!(
                same_length(placed.dx, band_start),
                "{categories} categories in a {frame_w}pt frame: `{category}`'s band \
                 starts at {band_start:.2}pt, got {:.2}pt in:\n{source}",
                placed.dx
            );
            // A height would hand Typst a second axis to centre in, moving the
            // baseline `line_category_baseline_pt` seats (issue #672).
            assert!(
                same_length(placed.height, 0.0),
                "a category label box states no height, got {:.2}pt in:\n{source}",
                placed.height
            );
        }
    }
}

/// The band is the layout, not a floor under it: a crowded axis gives each
/// label a band narrower than the retired 24pt box, and the label is laid out
/// in that narrower band — the same thing the column renderer does with `row`.
#[test]
fn a_crowded_line_axis_narrows_every_category_box_to_its_band() {
    let categories: usize = 16;
    let mut chart: Chart = dual_axis_line_chart();
    chart.categories = (1..=categories).map(|week| format!("W{week}")).collect();
    for series in &mut chart.series {
        series.values = (0..categories).map(|index| index as f64).collect();
    }

    let source: String = framed_chart_source(&chart, 400.0, 200.0);
    let (plot_x, _, plot_w, _) = plot_rect(&emitted_lines(&source));
    let band_w: f64 = plot_w / categories as f64;
    assert!(
        band_w < 24.0,
        "the fixture must crowd the axis below the retired box, got {band_w:.2}pt"
    );

    for (index, category) in chart.categories.iter().enumerate() {
        let placed: PlacedBox = placed_box_holding(&source, category);
        assert!(
            same_length(placed.width, band_w),
            "`{category}` must take its {band_w:.2}pt band rather than a floor, \
             got a {:.2}pt box in:\n{source}",
            placed.width
        );
        assert!(
            same_length(placed.dx, plot_x + index as f64 * band_w),
            "`{category}` must start on its own band boundary, got {:.2}pt in:\n{source}",
            placed.dx
        );
    }
}

/// A switched-off secondary axis prints no labels and reserves no gutter, but
/// its series still reads against its scale — hiding the axis is not the same
/// as putting the series on the primary one.
#[test]
fn a_deleted_secondary_axis_keeps_its_scale_but_prints_nothing() {
    let mut chart: Chart = dual_axis_line_chart();
    chart
        .secondary_value_axis
        .as_mut()
        .expect("the fixture carries a secondary axis")
        .deleted = true;
    let (frame_w, frame_h) = (400.0, 200.0);
    let source: String = framed_chart_source(&chart, frame_w, frame_h);
    let visible: String = framed_chart_source(&dual_axis_line_chart(), frame_w, frame_h);
    let (plot_x, plot_y, plot_w, plot_h) = plot_rect(&emitted_lines(&source));
    let (_, _, visible_plot_w, _) = plot_rect(&emitted_lines(&visible));

    let labels: Vec<(f64, String)> = emitted_value_labels(&source);
    assert!(
        labels.iter().all(|(dx, _)| *dx < plot_x),
        "a deleted secondary axis prints no label right of the plot in:\n{source}"
    );
    assert!(
        plot_w > visible_plot_w,
        "the plot reclaims the gutter a deleted axis would have used: {plot_w:.2}pt vs {visible_plot_w:.2}pt"
    );
    let acid: Vec<(f64, f64)> = emitted_path_points_of(&source, 1);
    let expected_acid_y: f64 = plot_y + (1.0 - 6.77 / 7.0) * plot_h;
    assert!(
        (acid[2].1 - expected_acid_y).abs() < 0.05,
        "the acid test still reads against its 0..7 axis at {expected_acid_y:.2}pt, got {:.2}pt",
        acid[2].1
    );
}

/// The budget columns with their spend line read against a secondary 0..200
/// axis: the classic column-and-line combo, where the line's own scale keeps
/// it legible over columns of a different magnitude (issue #1374).
fn combo_budget_chart_with_secondary_line() -> Chart {
    let mut chart: Chart = combo_budget_chart();
    chart.series[2].value_axis = crate::ir::ChartValueAxisRole::Secondary;
    chart.secondary_value_axis = Some(crate::ir::ChartSecondaryValueAxis {
        side: crate::ir::ValueAxisSide::Right,
        deleted: false,
        number_format: None,
        major_unit: Some(50.0),
        min: None,
        max: Some(200.0),
        major_tick_mark: AxisTickMark::None,
        line: crate::ir::ChartLine::Suppressed,
        text_font_family: None,
        text_style: crate::ir::ChartTextStyle::default(),
    });
    chart
}

/// A line laid over columns reads against the secondary axis its family
/// references, and that axis prints its own ticks right of the plot — while
/// the columns keep the primary scale their stacks set.
#[test]
fn a_line_over_columns_reads_against_its_secondary_axis() {
    let (frame_w, frame_h) = (400.0, 240.0);
    let source: String =
        framed_chart_source(&combo_budget_chart_with_secondary_line(), frame_w, frame_h);
    let (plot_x, plot_y, plot_w, plot_h) = plot_rect(&emitted_lines(&source));

    let points: Vec<(f64, f64)> = emitted_path_points(&source);
    assert_eq!(points.len(), 2, "two spend points in:\n{source}");
    let expected_y: f64 = plot_y + plot_h - (75.0 / 200.0) * plot_h;
    assert!(
        (points[1].1 - expected_y).abs() < 0.05,
        "75 on a 0..200 axis sits at {expected_y:.2}pt, got {:.2}pt in:\n{source}",
        points[1].1
    );

    let labels: Vec<(f64, String)> = emitted_value_labels(&source);
    let right: Vec<&str> = labels
        .iter()
        .filter(|(dx, _)| *dx >= plot_x + plot_w)
        .map(|(_, label)| label.as_str())
        .collect();
    assert_eq!(
        right,
        ["0", "50", "100", "150", "200"],
        "right labels in:\n{source}"
    );
    let primary_max: f64 = labels
        .iter()
        .filter(|(dx, _)| *dx < plot_x)
        .filter_map(|(_, label)| label.parse::<f64>().ok())
        .fold(0.0, f64::max);
    assert!(
        primary_max < 200.0,
        "the columns' axis stops at their tallest stack of 150, got {primary_max} in:\n{source}"
    );
    for label in ["0", "200"] {
        let placed: PlacedBox = placed_box_holding(&source, label);
        assert!(
            placed.dx + placed.width <= frame_w + 1e-6,
            "a secondary label box ends inside the {frame_w}pt frame, got {:.2}pt in:\n{source}",
            placed.dx + placed.width
        );
    }
}

// ---------------------------------------------------------------------------
// Automatic pie and doughnut plot inset (issue #1425)
// ---------------------------------------------------------------------------

/// A pie or doughnut of four equal slices, with no title and no legend.
///
/// Equal quarters put a wedge boundary at twelve, three, six and nine o'clock,
/// so the emitted vertices land on the circle's own extremes and the drawn
/// bounding box is the plot circle rather than a chord of it.
fn quartered_pie_chart(chart_type: ChartType) -> Chart {
    let mut chart = sized_bar_chart(CHART_DEFAULT_TEXT_PT);
    chart.chart_type = chart_type;
    chart.title = None;
    chart.has_legend = false;
    chart.host = crate::ir::ChartHost::Presentation;
    chart.categories = ["Q1", "Q2", "Q3", "Q4"].map(str::to_string).to_vec();
    chart.series.truncate(1);
    chart.series[0].values = vec![25.0, 25.0, 25.0, 25.0];
    chart.series[0].data_labels = DataLabels::default();
    chart
}

/// The centre, outer radius and inner radius of the drawn wedges, in points.
///
/// Every wedge vertex is stated as `((x pt, y pt), (in…), (out…))`, so the
/// first pair of each triple is a point on the drawn outline. The circle is
/// read back from those points alone, which keeps the assertion on what the
/// chart draws rather than on how the generator arrived at it.
/// Every `(Xpt, Ypt)` pair in `text`, in order.
fn coordinate_pairs(text: &str) -> Vec<(f64, f64)> {
    text.match_indices('(')
        .filter_map(|(start, _)| {
            let body: &str = text[start + 1..].split_once(')')?.0;
            let (x, y) = body.split_once(", ")?;
            Some((
                x.strip_suffix("pt")?.parse().ok()?,
                y.strip_suffix("pt")?.parse().ok()?,
            ))
        })
        .collect()
}

/// The vertices an emitted `curve(...)` passes through: its `curve.move`
/// start, then the end of every segment but the closing one back to it. Also
/// whether the first segment is straight, which marks a start vertex with no
/// handles, such as a pie wedge's centre.
fn curve_vertices(line: &str) -> (Vec<(f64, f64)>, bool) {
    let mut points: Vec<(f64, f64)> = Vec::new();
    let mut first_segment_is_straight: bool = false;
    for (index, segment) in line.split("curve.").skip(1).enumerate() {
        let pairs: Vec<(f64, f64)> = coordinate_pairs(segment);
        let end: Option<&(f64, f64)> = if segment.starts_with("cubic(") {
            pairs.get(2)
        } else if segment.starts_with("move(") || segment.starts_with("line(") {
            pairs.first()
        } else {
            None
        };
        if index == 1 {
            first_segment_is_straight = segment.starts_with("line(");
        }
        points.extend(end);
    }
    if points.len() > 1 && points.last() == points.first() {
        points.pop();
    }
    (points, first_segment_is_straight)
}

/// A wedge's outline vertices, without the bare centre a pie wedge starts at.
fn wedge_outline(line: &str) -> Vec<(f64, f64)> {
    let (mut points, starts_at_a_bare_centre) = curve_vertices(line);
    if starts_at_a_bare_centre && !points.is_empty() {
        points.remove(0);
    }
    points
}

fn emitted_ring_geometry(source: &str) -> ((f64, f64), f64, f64) {
    let mut points: Vec<(f64, f64)> = Vec::new();
    for line in source.lines().filter(|line| line.contains("curve(fill:")) {
        points.extend(wedge_outline(line));
    }
    assert!(!points.is_empty(), "no wedge path in:\n{source}");
    let min_x: f64 = points.iter().map(|(x, _)| *x).fold(f64::MAX, f64::min);
    let max_x: f64 = points.iter().map(|(x, _)| *x).fold(f64::MIN, f64::max);
    let min_y: f64 = points.iter().map(|(_, y)| *y).fold(f64::MAX, f64::min);
    let max_y: f64 = points.iter().map(|(_, y)| *y).fold(f64::MIN, f64::max);
    let centre: (f64, f64) = ((min_x + max_x) / 2.0, (min_y + max_y) / 2.0);
    let radius = |(x, y): &(f64, f64)| ((x - centre.0).powi(2) + (y - centre.1).powi(2)).sqrt();
    let outer: f64 = points.iter().map(radius).fold(f64::MIN, f64::max);
    let inner: f64 = points.iter().map(radius).fold(f64::MAX, f64::min);
    (centre, outer, inner)
}

#[test]
fn an_automatic_powerpoint_pie_reserves_the_native_plot_inset() {
    // Native PowerPoint for Mac 16.112 exports, measured off the wedge paths
    // with `scripts/probe_harness.py --backend office`: the plot circle is the
    // chart area's smaller side less a constant 22pt, whatever that side is.
    for (frame_w, frame_h) in [
        (480.0, 320.0),
        (340.0, 340.0),
        (563.125, 340.125),
        (240.0, 320.0),
        (208.8, 122.4),
    ] {
        let source: String =
            framed_chart_source(&quartered_pie_chart(ChartType::Pie), frame_w, frame_h);
        let (centre, outer, _) = emitted_ring_geometry(&source);
        assert!(
            (outer * 2.0 - (frame_w.min(frame_h) - 22.0)).abs() < 0.01,
            "{frame_w}x{frame_h}pt frame: expected a {:.3}pt circle, drew {:.3}pt",
            frame_w.min(frame_h) - 22.0,
            outer * 2.0
        );
        assert!(
            (centre.0 - frame_w / 2.0).abs() < 0.01 && (centre.1 - frame_h / 2.0).abs() < 0.01,
            "{frame_w}x{frame_h}pt frame: the inset is symmetric, so the circle stays centred; \
             got {centre:?}"
        );
    }
}

#[test]
fn an_automatic_powerpoint_doughnut_insets_its_ring_and_keeps_its_hole() {
    for (frame_w, frame_h, hole) in [
        (563.125, 340.125, 71_u32),
        (208.8, 122.4, 62),
        (320.0, 480.0, 40),
    ] {
        let mut chart = quartered_pie_chart(ChartType::Doughnut);
        chart.hole_size_percent = Some(hole);
        let source: String = framed_chart_source(&chart, frame_w, frame_h);
        let (centre, outer, inner) = emitted_ring_geometry(&source);
        assert!(
            (outer * 2.0 - (frame_w.min(frame_h) - 22.0)).abs() < 0.01,
            "{frame_w}x{frame_h}pt frame: expected a {:.3}pt ring, drew {:.3}pt",
            frame_w.min(frame_h) - 22.0,
            outer * 2.0
        );
        assert!(
            (inner / outer - f64::from(hole) / 100.0).abs() < 1e-6,
            "the inset must not disturb the authored {hole}% hole, got {:.4}",
            inner / outer
        );
        assert!(
            (centre.0 - frame_w / 2.0).abs() < 0.01 && (centre.1 - frame_h / 2.0).abs() < 0.01,
            "{frame_w}x{frame_h}pt frame: ring centre {centre:?}"
        );
    }
}

#[test]
fn a_worksheet_pie_keeps_the_plot_size_excel_was_measured_at() {
    // The 22pt inset is a PowerPoint measurement. Excel's own automatic pie
    // inset is unmeasured, so a worksheet chart keeps what it drew before.
    for host in [
        crate::ir::ChartHost::Spreadsheet,
        crate::ir::ChartHost::SpreadsheetChartsheet,
        crate::ir::ChartHost::WordProcessing,
    ] {
        let mut chart = quartered_pie_chart(ChartType::Pie);
        chart.host = host;
        let source: String = framed_chart_source(&chart, 480.0, 320.0);
        let (_, outer, _) = emitted_ring_geometry(&source);
        assert!(
            (outer * 2.0 - 320.0).abs() < 0.01,
            "{host:?}: expected the full 320pt, drew {:.3}pt",
            outer * 2.0
        );
    }
}

#[test]
fn a_stated_plot_rectangle_keeps_a_pie_off_the_automatic_inset() {
    // `c:plotArea/c:layout` places the plot itself, so the automatic inset is
    // not the rule that applies; the stated case is left exactly as it was.
    let mut chart = quartered_pie_chart(ChartType::Pie);
    chart.plot_area_layout = Some(crate::ir::ChartPlotAreaLayout {
        x: 0.05,
        y: 0.05,
        width: 0.9,
        height: 0.9,
    });
    let source: String = framed_chart_source(&chart, 480.0, 320.0);
    let (_, outer, _) = emitted_ring_geometry(&source);
    assert!(
        (outer * 2.0 - 320.0).abs() < 0.01,
        "expected the stated-layout pie unchanged at 320pt, drew {:.3}pt",
        outer * 2.0
    );
}

// ---------------------------------------------------------------------------
// Authored first-slice rotation (issue #1429)
// ---------------------------------------------------------------------------

/// Every drawn wedge's outline vertices, one entry per wedge, in the order the
/// wedges are emitted.
///
/// Each vertex is stated as `((x pt, y pt), (in…), (out…))`, so splitting on
/// the opening `((` yields the outline and skips the bare centre point a pie
/// wedge starts from.
fn emitted_wedge_outlines(source: &str) -> Vec<Vec<(f64, f64)>> {
    let mut wedges: Vec<Vec<(f64, f64)>> = Vec::new();
    for line in source.lines().filter(|line| line.contains("curve(fill:")) {
        let outline: Vec<(f64, f64)> = wedge_outline(line);
        if !outline.is_empty() {
            wedges.push(outline);
        }
    }
    wedges
}

/// `angle` measured clockwise from twelve o'clock about `centre`, in degrees.
///
/// Typst places from the top-left with y growing downward, so a clockwise turn
/// on the page runs from `-y` towards `+x`.
fn clockwise_degrees_from_twelve(centre: (f64, f64), point: (f64, f64)) -> f64 {
    (point.0 - centre.0)
        .atan2(-(point.1 - centre.1))
        .to_degrees()
        .rem_euclid(360.0)
}

/// Signed difference between two clockwise angles, folded into ±180 degrees so
/// a boundary at 359.99 and one at 0.01 read as neighbours.
fn degrees_apart(left: f64, right: f64) -> f64 {
    (left - right + 180.0).rem_euclid(360.0) - 180.0
}

/// A four-point pie whose shares are the ones `GENERAL SERVICES.pptx` page 13
/// plots, so no two boundaries coincide and a rotation cannot hide in a
/// symmetry.
fn unequal_pie_chart(chart_type: ChartType, first_slice_angle_deg: Option<u32>) -> Chart {
    let mut chart = quartered_pie_chart(chart_type);
    chart.series[0].values = vec![39.0, 20.0, 20.0, 21.0];
    chart.hole_size_percent = Some(71);
    chart.first_slice_angle_deg = first_slice_angle_deg;
    chart
}

#[test]
fn an_authored_first_slice_angle_turns_the_whole_plot_clockwise() {
    // Native PowerPoint for Mac 16.112 exports of `GENERAL SERVICES.pptx`,
    // read off the exported wedge paths with `mutool draw -F trace`: page 13
    // declares `<c:firstSliceAng val="12"/>` and starts its first boundary
    // 11.996 degrees clockwise of twelve o'clock, page 14 declares 11 and
    // starts at 11.003. Every following boundary is that start plus the
    // shares already plotted, so the rotation moves the plot without
    // resizing a wedge (issue #1429).
    const FRAME_W: f64 = 563.125;
    const FRAME_H: f64 = 340.125;
    let centre: (f64, f64) = (FRAME_W / 2.0, FRAME_H / 2.0);
    let shares: [f64; 4] = [39.0, 20.0, 20.0, 21.0];

    for chart_type in [ChartType::Pie, ChartType::Doughnut] {
        for authored in [0_u32, 11, 12, 90, 137, 270, 360] {
            let chart: Chart = unequal_pie_chart(chart_type.clone(), Some(authored));
            let source: String = framed_chart_source(&chart, FRAME_W, FRAME_H);
            let wedges: Vec<Vec<(f64, f64)>> = emitted_wedge_outlines(&source);
            assert_eq!(
                wedges.len(),
                shares.len(),
                "{chart_type:?} at {authored} degrees draws one path per share, got:\n{source}"
            );

            let mut expected: f64 = f64::from(authored);
            for (index, share) in shares.iter().enumerate() {
                let start: f64 = clockwise_degrees_from_twelve(centre, wedges[index][0]);
                assert!(
                    degrees_apart(start, expected).abs() < 0.01,
                    "{chart_type:?} at {authored} degrees: wedge {index} starts at \
                     {start:.3} degrees, expected {:.3}",
                    expected.rem_euclid(360.0)
                );
                expected += share / 100.0 * 360.0;
            }
        }
    }
}

/// Triangulation on the absent case: `ST_FirstSliceAng` defaults to 0, so a
/// chart that declares nothing keeps the twelve o'clock start.
#[test]
fn a_chart_without_a_first_slice_angle_starts_at_twelve_oclock() {
    const FRAME_W: f64 = 480.0;
    const FRAME_H: f64 = 320.0;
    let centre: (f64, f64) = (FRAME_W / 2.0, FRAME_H / 2.0);

    for chart_type in [ChartType::Pie, ChartType::Doughnut] {
        let chart: Chart = unequal_pie_chart(chart_type.clone(), None);
        let source: String = framed_chart_source(&chart, FRAME_W, FRAME_H);
        let wedges: Vec<Vec<(f64, f64)>> = emitted_wedge_outlines(&source);
        let start: f64 = clockwise_degrees_from_twelve(centre, wedges[0][0]);
        assert!(
            degrees_apart(start, 0.0).abs() < 0.01,
            "{chart_type:?} with no authored angle starts at {start:.3} degrees"
        );
    }
}

/// The rotation turns the plot; it must not resize it. The ring keeps the
/// automatic plot diameter of #1425, the authored hole ratio, and its centre.
#[test]
fn a_turned_ring_keeps_its_centre_diameter_and_hole() {
    const FRAME_W: f64 = 563.125;
    const FRAME_H: f64 = 340.125;
    let centre: (f64, f64) = (FRAME_W / 2.0, FRAME_H / 2.0);
    let expected_outer: f64 = (FRAME_H - 22.0) / 2.0;

    for authored in [None, Some(0), Some(12), Some(137), Some(359)] {
        let chart: Chart = unequal_pie_chart(ChartType::Doughnut, authored);
        let source: String = framed_chart_source(&chart, FRAME_W, FRAME_H);
        let radii: Vec<f64> = emitted_wedge_outlines(&source)
            .iter()
            .flatten()
            .map(|(x, y)| ((x - centre.0).powi(2) + (y - centre.1).powi(2)).sqrt())
            .collect();
        let outer: f64 = radii.iter().copied().fold(f64::MIN, f64::max);
        let inner: f64 = radii.iter().copied().fold(f64::MAX, f64::min);
        assert!(
            (outer - expected_outer).abs() < 0.01,
            "{authored:?}: outer radius {outer:.3}pt, expected {expected_outer:.3}pt"
        );
        assert!(
            (inner / outer - 0.71).abs() < 1e-6,
            "{authored:?}: hole ratio {:.4}, expected 0.7100",
            inner / outer
        );
    }
}
