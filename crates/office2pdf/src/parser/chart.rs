//! Shared chart XML parser for DOCX, PPTX and XLSX packages.
//!
//! Parses chart*.xml parts and extracts chart type,
//! title, category labels, and series data into IR `Chart` structs.

use crate::parser::xml_util::OOXML_XML_VERSION;
use quick_xml::Reader;
use quick_xml::events::Event;

use super::drawingml::{self, SchemeColors};
use super::xml_util;
use crate::ir::{
    AxisTickMark, BarBandLayout, Chart, ChartAreaFill, ChartAreaOutline, ChartFillMode,
    ChartGrouping, ChartHost, ChartLine, ChartMarkerStyle, ChartPlotAreaLayout,
    ChartSecondaryValueAxis, ChartSeries, ChartTextStyle, ChartTitleLayout, ChartType,
    ChartValueAxisRole, Color, DataLabelPosition, DataLabels, LegendPosition, MarkerSymbol,
    ValueAxisSide,
};

/// Mapping from XML chart element tag names to their corresponding `ChartType`.
/// Both 2-D and 3-D variants map to the same logical type.
///
/// The bar family maps to `Column` because that is the orientation ECMA-376
/// gives `ST_BarDir`'s `val` attribute by default, and the one Excel and
/// PowerPoint write for their default clustered chart. `<c:barDir val="bar"/>`
/// overrides it; see [`bar_direction_chart_type`].
const CHART_TAG_TYPES: &[(&[u8], ChartType)] = &[
    (b"barChart", ChartType::Column),
    (b"bar3DChart", ChartType::Column),
    (b"lineChart", ChartType::Line),
    (b"line3DChart", ChartType::Line),
    (b"pieChart", ChartType::Pie),
    (b"pie3DChart", ChartType::Pie),
    (b"doughnutChart", ChartType::Doughnut),
    (b"areaChart", ChartType::Area),
    (b"scatterChart", ChartType::Scatter),
];

/// Display labels for the plot-area families ECMA-376 defines that no plot
/// implementation covers. They render through the data-table fallback, which
/// prints this label as the chart's kind.
///
/// Dropping them instead made the chart vanish with no diagnostic, taking the
/// whole graphic frame with it (issue #544).
const UNPLOTTED_CHART_LABELS: &[(&[u8], &str)] = &[
    (b"radarChart", crate::ir::RADAR_CHART_LABEL),
    (b"bubbleChart", "Bubble Chart"),
    (b"stockChart", "Stock Chart"),
    (b"surfaceChart", "Surface Chart"),
    (b"surface3DChart", "Surface Chart"),
    (b"area3DChart", "Area Chart"),
];

/// `<c:ofPieChart>` covers two shapes, told apart by `<c:ofPieType>`. ECMA-376
/// defaults `ST_OfPieType` to `pie`.
fn of_pie_label(of_pie_type: Option<&str>) -> &'static str {
    match of_pie_type {
        Some("bar") => "Bar of Pie Chart",
        _ => "Pie of Pie Chart",
    }
}

/// Resolve a plot-area element to the chart type it opens.
///
/// Families with a plot implementation take it; the rest keep their data and
/// a readable label so the fallback can draw them. Anything else whose name
/// ends in `Chart` is a family this list has not caught up with, and is better
/// rendered as a table than dropped.
fn chart_type_for_tag(tag: &[u8]) -> Option<ChartType> {
    if let Some((_, chart_type)) = CHART_TAG_TYPES.iter().find(|(name, _)| *name == tag) {
        return Some(chart_type.clone());
    }
    if let Some((_, label)) = UNPLOTTED_CHART_LABELS.iter().find(|(name, _)| *name == tag) {
        return Some(ChartType::Other(label.to_string()));
    }
    // The suffix is enough on its own: the part's own `chartSpace` and `chart`
    // elements differ in case or ending, and every element that does carry it
    // is a plot-area family.
    let name: &str = std::str::from_utf8(tag).ok()?;
    name.ends_with("Chart")
        .then(|| ChartType::Other(generic_chart_label(name)))
}

/// Turn an unknown `fooBarChart` element name into `Foo Bar Chart`.
fn generic_chart_label(element: &str) -> String {
    let mut label = String::new();
    for (index, character) in element.char_indices() {
        if index > 0 && character.is_ascii_uppercase() {
            label.push(' ');
        }
        if index == 0 {
            label.extend(character.to_uppercase());
        } else {
            label.push(character);
        }
    }
    label
}

/// Resolve a `<c:legendPos>` value to the edge the legend sits on.
fn legend_position_for(value: &str) -> LegendPosition {
    match value {
        "b" => LegendPosition::Bottom,
        "l" => LegendPosition::Left,
        "t" => LegendPosition::Top,
        "tr" => LegendPosition::TopRight,
        _ => LegendPosition::Right,
    }
}

/// Resolve a `<c:grouping>` value to the way a category's series combine.
///
/// Line and area charts spell their unstacked form `standard` where the bar
/// family spells it `clustered`; both mean one mark per series.
fn chart_grouping_for(value: &str) -> ChartGrouping {
    match value {
        "stacked" => ChartGrouping::Stacked,
        "percentStacked" => ChartGrouping::PercentStacked,
        _ => ChartGrouping::Clustered,
    }
}

/// Resolve a `<c:barDir>` value to the chart orientation it selects.
///
/// `<c:barDir>` is exclusive to the bar family, so `None` — either an absent
/// element or a non-bar chart — leaves the tag's own mapping in place.
fn bar_direction_chart_type(direction: Option<&str>) -> Option<ChartType> {
    match direction? {
        "bar" => Some(ChartType::Bar),
        "col" => Some(ChartType::Column),
        _ => None,
    }
}

/// Parse a chart XML file (e.g., `word/charts/chart1.xml`) into a `Chart` IR.
pub(crate) fn parse_chart_xml(xml: &str, scheme: &SchemeColors<'_>) -> Option<Chart> {
    parse_chart_xml_with_stroke_defaults(xml, scheme, crate::ir::ChartStrokeGeometry::default())
}

pub(super) fn parse_chart_xml_with_stroke_defaults(
    xml: &str,
    scheme: &SchemeColors<'_>,
    stroke_defaults: crate::ir::ChartStrokeGeometry,
) -> Option<Chart> {
    let mut reader = Reader::from_str(xml);
    let mut chart_type = None;
    let mut hole_size_percent: Option<u32> = None;
    let mut first_slice_angle_deg: Option<u32> = None;
    let mut title = None;
    let mut categories: Vec<String> = Vec::new();
    let mut series: Vec<ChartSeries> = Vec::new();
    let mut grouping: Option<ChartGrouping> = None;
    let mut gap_width_percent: Option<f64> = None;
    let mut overlap_percent: Option<f64> = None;
    let mut legend_position: Option<LegendPosition> = None;
    // `<c:legend>` presence, not its position: a chart that declares no legend
    // still gets a default position, so the position alone cannot say whether
    // one was asked for (issue #762).
    let mut has_legend: bool = false;
    let mut auto_title_deleted: bool = false;
    // `<c:title>` present but naming no text of its own. Office supplies the
    // string for one of those, so the element's absence and its emptiness are
    // different states and `title` alone cannot tell them apart (issue #1146).
    let mut has_automatic_title: bool = false;
    // Every axis the plot area declares, in document order. A combo plot area
    // may carry two of each — one pair per axis group — and which family reads
    // against which is settled by `<c:axId>` reference once the whole plot
    // area has been read (issue #1374).
    let mut category_axes: Vec<Axis> = Vec::new();
    let mut value_axes: Vec<Axis> = Vec::new();
    // The series each family declared, with the `<c:axId>` pair it names.
    let mut family_axis_ids: Vec<(std::ops::Range<usize>, Vec<String>)> = Vec::new();
    // `c:chartSpace/c:spPr` is a *sibling* of `c:chart`, and the schema puts it
    // after it. This loop is flat over every Start event, so without that
    // marker a `c:spPr` belonging to `c:plotArea` would be read as the chart
    // area's own (#637).
    let mut chart_element_ended: bool = false;
    let mut chart_area_fill: ChartAreaFill = ChartAreaFill::Unspecified;
    let mut chart_area_outline: ChartAreaOutline = ChartAreaOutline::Default;
    // `c:chartSpace/c:txPr` is a sibling of `c:chart` for the same reason
    // `c:spPr` is, so it needs the same marker: an axis carries a `c:txPr` of
    // its own inside `c:plotArea`, and a flat loop would read that one (#668).
    let mut text_font_family: Option<String> = None;
    let mut text_style: ChartTextStyle = ChartTextStyle::default();
    // The title's own `c:tx/c:rich` run properties over its `c:txPr` govern the
    // title alone, the rich text winning where both speak (issue #1424); the
    // chart space's `c:txPr` governs everything else and is a poor stand-in
    // for it (issue #1215).
    let mut title_text_style: ChartTextStyle = ChartTextStyle::default();
    let mut title_layout: Option<ChartTitleLayout> = None;
    // A legend can override the chart space's run properties independently of
    // its position and visibility (issue #1236).
    let mut legend_text_style: ChartTextStyle = ChartTextStyle::default();
    // `c:layout` is written by the title, the legend and every data-label group
    // as well, so the plot area's own is told from theirs by where it sits. The
    // elements that carry the others consume their own subtrees before this
    // loop sees them, but the marker keeps that from being an assumption (#1182).
    let mut in_plot_area: bool = false;
    let mut plot_area_layout: Option<ChartPlotAreaLayout> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                let tag: &[u8] = local.as_ref();
                if tag == b"spPr" && chart_element_ended {
                    (chart_area_fill, chart_area_outline) =
                        parse_chart_area_properties(&mut reader, scheme);
                } else if tag == b"txPr" && chart_element_ended {
                    let (family, style, _) = parse_chart_text_properties(&mut reader, scheme);
                    text_font_family = family;
                    text_style = style;
                } else if tag == b"plotArea" {
                    in_plot_area = true;
                } else if tag == b"layout" && in_plot_area && plot_area_layout.is_none() {
                    plot_area_layout = parse_plot_area_layout(&mut reader);
                } else if tag == b"legend" {
                    // Declared, unless the element switches itself off.
                    let (deleted, position, style) = parse_legend(&mut reader, scheme);
                    has_legend = !deleted;
                    legend_position = legend_position.or(position);
                    legend_text_style = style;
                } else if tag == b"title" && title.is_none() {
                    let (text, names_own_text, style, layout) =
                        parse_chart_title(&mut reader, scheme);
                    title = text;
                    has_automatic_title = !names_own_text;
                    title_text_style = style;
                    title_layout = layout;
                } else if tag == b"catAx" {
                    category_axes.push(parse_axis(&mut reader, b"catAx", scheme));
                } else if tag == b"valAx" {
                    value_axes.push(parse_axis(&mut reader, b"valAx", scheme));
                } else if let Some(ct) = chart_type_for_tag(tag) {
                    let mut plot: PlotAreaProps = PlotAreaProps::default();
                    let family_first_series: usize = series.len();
                    parse_chart_series(
                        &mut reader,
                        tag,
                        &mut categories,
                        &mut series,
                        &mut plot,
                        scheme,
                    );
                    family_axis_ids
                        .push((family_first_series..series.len(), plot.axis_ids.clone()));
                    let family: ChartType = match ct {
                        // `<c:ofPieChart>` names its shape in a child element,
                        // so the label waits until the body has been read.
                        ChartType::Other(_) if tag == b"ofPieChart" => {
                            ChartType::Other(of_pie_label(plot.of_pie_type.as_deref()).to_string())
                        }
                        other => {
                            bar_direction_chart_type(plot.bar_direction.as_deref()).unwrap_or(other)
                        }
                    };
                    // Which family drew which series, so a combo plot area's
                    // line does not draw as a column and vice versa. Normalised
                    // against the chart's own family once it is settled, below
                    // (issue #1067).
                    // dataPointLine is for 2-D line/scatter and unfilled radar.
                    // line3DChart and filled radar normalize to the same IR
                    // families but use different chart-style entries.
                    let uses_line_defaults = matches!(tag, b"lineChart" | b"scatterChart")
                        || (tag == b"radarChart" && plot.radar_style.as_deref() != Some("filled"));
                    for entry in &mut series[family_first_series..] {
                        entry.plot_type = Some(family.clone());
                        if uses_line_defaults {
                            let geometry = &mut entry.line_geometry;
                            geometry.cap = geometry.cap.or(stroke_defaults.cap);
                            if geometry.join.is_none() {
                                geometry.join = stroke_defaults.join;
                                geometry.miter_limit = stroke_defaults.miter_limit;
                            }
                        }
                    }
                    // The bar family governs the chart: the value scale and the
                    // category bands are the ones its columns are drawn to, and
                    // the other families lay over them. Without this a
                    // `<c:lineChart>` following a `<c:barChart>` took the type
                    // and the grouping with it, and the columns vanished
                    // (issue #1067).
                    //
                    // A scatter family yields to any family already governing,
                    // for the same reason at the other end: `<c:scatterChart>`
                    // carries no `<c:cat>` — each point states its own x — so it
                    // never declares the category bands the family beside it
                    // did. Letting the `<c:scatterChart>` that follows the
                    // `<c:lineChart>` in `Monthly college budget1.xlsx` take the
                    // type dropped the whole chart to the data-table fallback,
                    // which no scatter plot exists to spare it (issue #1123).
                    let yields_to_governing_family: bool =
                        chart_type.is_some() && matches!(family, ChartType::Scatter);
                    if !matches!(chart_type, Some(ChartType::Bar | ChartType::Column))
                        && !yields_to_governing_family
                    {
                        chart_type = Some(family);
                        grouping = plot.grouping.as_deref().map(chart_grouping_for);
                        // Only the doughnut family writes it.
                        if matches!(chart_type, Some(ChartType::Doughnut)) {
                            hole_size_percent =
                                plot.hole_size.as_deref().and_then(|v| v.parse().ok());
                        }
                        // Both circular families write it; no other family
                        // declares the element at all.
                        if matches!(chart_type, Some(ChartType::Pie | ChartType::Doughnut)) {
                            first_slice_angle_deg = plot
                                .first_slice_ang
                                .as_deref()
                                .and_then(first_slice_angle_degrees);
                        }
                    }
                    // A combo plot area holds one element per chart family and
                    // only the bar family carries these two, so the family that
                    // declared them keeps them however many follow it.
                    let (gap_width, overlap) = plot.bar_band_layout();
                    gap_width_percent = gap_width_percent.or(gap_width);
                    overlap_percent = overlap_percent.or(overlap);
                }
            }
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"legend" => {
                has_legend = true;
            }
            // `<c:title/>` with no children names no text either, which is the
            // whole condition for an automatic title.
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"title" && title.is_none() => {
                has_automatic_title = true;
            }
            // `CT_Boolean`, so always self-closing; an earlier arm already
            // owns every `Start` event in this loop.
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"autoTitleDeleted" => {
                auto_title_deleted = ct_boolean(e);
            }
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"legendPos" => {
                legend_position = xml_util::get_attr_str(e, b"val")
                    .as_deref()
                    .map(legend_position_for);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"plotArea" => {
                in_plot_area = false;
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"chart" => {
                chart_element_ended = true;
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    let chart_type = chart_type?;
    let default_band_layout: BarBandLayout = BarBandLayout::default();

    // `plot_type` records the family only where it differs from the chart's,
    // so a single-family chart carries none of it and hand-built IR reads the
    // same as parsed IR (issue #1067).
    for entry in &mut series {
        if entry.plot_type.as_ref() == Some(&chart_type) {
            entry.plot_type = None;
        }
    }

    let AxisGroups {
        category_axis,
        value_axis,
        secondary_value_axis,
    } = resolve_axis_groups(category_axes, value_axes, &family_axis_ids, &mut series);

    // Charts may omit <c:cat> entirely; Excel then labels the category axis
    // 1..N (the point count of the longest series).
    if categories.is_empty() {
        let point_count: usize = series.iter().map(|s| s.values.len()).max().unwrap_or(0);
        categories = (1..=point_count).map(|i| i.to_string()).collect();
    }

    // A label position the part left out is settled by the grouping, which is
    // only known once the whole plot has been read: a clustered bar puts its
    // labels beyond the bar's end, a stacked one centres them because an
    // outside label would land on the segment above (ECMA-376 §21.2.2.49).
    let grouping: ChartGrouping = grouping.unwrap_or_default();
    if matches!(grouping, ChartGrouping::Clustered) {
        for entry in &mut series {
            if !entry.data_labels.position_stated {
                entry.data_labels.position = DataLabelPosition::OutsideEnd;
            }
        }
    }

    Some(Chart {
        chart_type,
        hole_size_percent,
        first_slice_angle_deg,
        title,
        categories,
        series,
        grouping,
        legend_position: legend_position.unwrap_or_default(),
        has_legend,
        auto_title_deleted,
        has_automatic_title,
        title_layout,
        plot_area_layout,
        // The chart part names the drawing its user shapes live in through a
        // relationship, which only the package that holds it can resolve — as
        // with the theme above, the loader fills these in (issue #1186).
        user_shapes: Vec::new(),
        category_axis_title: category_axis.title,
        value_axis_title: value_axis.title,
        category_axis_major_tick_mark: category_axis.major_tick_mark,
        value_axis_major_tick_mark: value_axis.major_tick_mark,
        category_axis_line: category_axis.line,
        value_axis_line: value_axis.line,
        value_axis_major_unit: value_axis.major_unit,
        value_axis_min: value_axis.min,
        value_axis_max: value_axis.max,
        // Office hangs the gridlines off whichever axis they run across; the
        // value axis carries the horizontal set our renderer draws.
        // An absent element draws no gridlines. A present but unstyled
        // element keeps its automatic stroke rather than borrowing another
        // axis' styling (issue #1271).
        major_gridline_line: value_axis
            .gridline
            .or(category_axis.gridline)
            .unwrap_or(ChartLine::Suppressed),
        category_axis_deleted: category_axis.deleted,
        value_axis_deleted: value_axis.deleted,
        bar_band_layout: BarBandLayout {
            gap_width_percent: gap_width_percent.unwrap_or(default_band_layout.gap_width_percent),
            overlap_percent: overlap_percent.unwrap_or(default_band_layout.overlap_percent),
        },
        // The chart part names no theme of its own; the package that holds it
        // does. Whoever loaded this XML fills these in, since only they know
        // which theme part applies.
        theme_accent_colors: Vec::new(),
        chart_area_fill,
        chart_area_outline,
        // As with the theme, the chart part does not know which application's
        // package holds it; the loader sets this (issue #823).
        host: ChartHost::default(),
        // A `+mn-lt` token stays as written for the same reason: resolving it
        // needs the package theme, which only the loader has.
        text_font_family,
        text_style,
        title_text_style,
        legend_text_style,
        category_axis_text_font_family: category_axis.font_family,
        value_axis_text_font_family: value_axis.font_family,
        category_axis_text_style: category_axis.text_style,
        value_axis_text_style: value_axis.text_style,
        category_axis_number_format: category_axis.number_format,
        value_axis_number_format: value_axis.number_format,
        secondary_value_axis: secondary_value_axis.map(|axis| ChartSecondaryValueAxis {
            side: match axis.position.as_deref() {
                Some("l") => ValueAxisSide::Left,
                _ => ValueAxisSide::Right,
            },
            deleted: axis.deleted,
            number_format: axis.number_format,
            major_unit: axis.major_unit,
            min: axis.min,
            max: axis.max,
            major_tick_mark: axis.major_tick_mark,
            line: axis.line,
            text_font_family: axis.font_family,
            text_style: axis.text_style,
        }),
    })
}

/// Read `c:plotArea/c:layout` into the rectangle it gives the plotting area.
///
/// Only the shape Excel and PowerPoint write for a plot the user dragged or
/// resized is taken: `layoutTarget="inner"`, both modes `edge`, and all four
/// values present. `factor` — the `ST_LayoutMode` default, and so what an
/// omitted mode means — offsets the automatic layout instead of naming an edge,
/// and `outer` measures the plot area with its tick labels rather than the
/// plotting rectangle. Reading either as an edge fraction would move the plot
/// somewhere nothing asked for, so anything else keeps the automatic layout
/// (issue #1182).
fn parse_plot_area_layout(reader: &mut Reader<&[u8]>) -> Option<ChartPlotAreaLayout> {
    let mut layout_target: Option<String> = None;
    let mut x_mode: Option<String> = None;
    let mut y_mode: Option<String> = None;
    let mut x: Option<f64> = None;
    let mut y: Option<f64> = None;
    let mut width: Option<f64> = None;
    let mut height: Option<f64> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                let value = || xml_util::get_attr_str(e, b"val");
                let number = || value().and_then(|raw| raw.parse::<f64>().ok());
                match local.as_ref() {
                    b"layoutTarget" => layout_target = value(),
                    b"xMode" => x_mode = value(),
                    b"yMode" => y_mode = value(),
                    b"x" => x = number(),
                    b"y" => y = number(),
                    b"w" => width = number(),
                    b"h" => height = number(),
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"layout" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if layout_target.as_deref() != Some("inner")
        || x_mode.as_deref() != Some("edge")
        || y_mode.as_deref() != Some("edge")
    {
        return None;
    }
    let (x, y, width, height) = (x?, y?, width?, height?);
    // A plot with no extent draws nothing, and a non-finite fraction would take
    // the whole placement with it.
    if !(x.is_finite() && y.is_finite() && width > 0.0 && height > 0.0) {
        return None;
    }
    // A rectangle that starts before the chart area or is larger than it is
    // dropped whole: native Excel for Mac 16 exports of
    // `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx` with one value of the
    // cash-flow chart's layout rewritten — `c:x` -0.1, `c:y` -0.1, `c:w` 1.05
    // or 1.2, `c:h` 1.05 — each laid that plot out automatically, the same
    // inset on all four sides with the stated `c:x` and `c:y` ignored too,
    // where `c:x` 0 and `c:w` 1 both still counted (issue #1272). A rectangle
    // that merely runs past the far edge is kept; the renderer pulls it back.
    if x < 0.0 || y < 0.0 || width > 1.0 || height > 1.0 {
        return None;
    }
    Some(ChartPlotAreaLayout {
        x,
        y,
        width,
        height,
    })
}

/// Read `c:title/c:layout` into its chart-relative top-left anchor.
///
/// Unlike a plot-area rectangle, a PowerPoint title commonly states only
/// edge-mode `x` and `y`: its width and height remain automatic. An omitted
/// mode means `factor`, whose values offset an application-computed position,
/// so only two explicit edge modes are deterministic enough to carry into the
/// renderer (issue #1423).
fn parse_title_layout(reader: &mut Reader<&[u8]>) -> Option<ChartTitleLayout> {
    let mut x_mode: Option<String> = None;
    let mut y_mode: Option<String> = None;
    let mut x: Option<f64> = None;
    let mut y: Option<f64> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let local = e.local_name();
                let value = || xml_util::get_attr_str(e, b"val");
                let number = || value().and_then(|raw| raw.parse::<f64>().ok());
                match local.as_ref() {
                    b"xMode" => x_mode = value(),
                    b"yMode" => y_mode = value(),
                    b"x" => x = number(),
                    b"y" => y = number(),
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"layout" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if x_mode.as_deref() != Some("edge") || y_mode.as_deref() != Some("edge") {
        return None;
    }
    let (x, y) = (x?, y?);
    (x.is_finite() && y.is_finite()).then_some(ChartTitleLayout { x, y })
}

/// Read `c:chartSpace/c:txPr` into the face and the run properties its
/// `a:defRPr` declares.
///
/// `a:latin`'s absence is the common case — the chart then inherits the theme's
/// minor font — and is reported as `None` rather than as a face, so the loader
/// can tell "said nothing" from "said this" (issue #668). The same holds for
/// each run property (issue #669).
fn parse_chart_text_properties(
    reader: &mut Reader<&[u8]>,
    scheme: &SchemeColors<'_>,
) -> (Option<String>, ChartTextStyle, bool) {
    let mut typeface: Option<String> = None;
    let mut style: ChartTextStyle = ChartTextStyle::default();
    let mut ellipsis: bool = false;
    // `a:solidFill` also appears under `a:ln` and other siblings inside a
    // `c:txPr`, so the colour is only taken while inside the run properties
    // themselves (issue #916).
    let mut in_def_rpr: bool = false;
    let mut in_solid_fill: bool = false;
    loop {
        match reader.read_event() {
            // A colour element written with children keeps its transforms
            // there, so its start tag alone is not the colour: the
            // `<a:schemeClr val="tx1">` every string of the workbook in #1160
            // names is lifted by lumMod 65% / lumOff 35% to #595959, and
            // reading only the tag printed all of them pure black. This has to
            // precede the shared start/empty arm below, which cannot consume
            // the children (issue #1160).
            Ok(Event::Start(ref e))
                if in_solid_fill
                    && matches!(
                        e.local_name().as_ref(),
                        b"srgbClr" | b"schemeClr" | b"sysClr"
                    ) =>
            {
                // Parsed even once a colour is known, because the element has
                // to be consumed either way to leave the reader on its sibling
                // — `<a:latin>` follows `<a:solidFill>` inside `<a:defRPr>`.
                let parsed = drawingml::parse_color_from_start(reader, e, scheme);
                // `<a:alpha>` is this colour's own child, so it is taken with
                // the colour it modifies and never on its own (issue #1677).
                if style.color.is_none() && parsed.color.is_some() {
                    style.color = parsed.color;
                    style.alpha = parsed.alpha;
                }
            }
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                match e.local_name().as_ref() {
                    b"bodyPr" => {
                        ellipsis = xml_util::get_attr_str(e, b"vertOverflow")
                            .is_some_and(|value| value == "ellipsis");
                    }
                    b"latin" if typeface.is_none() => {
                        // An empty `typeface=""` names no face; Office writes
                        // that to mean "inherit", which is what `None` says.
                        typeface = xml_util::get_attr_str(e, b"typeface")
                            .filter(|face| !face.trim().is_empty());
                    }
                    b"defRPr" => {
                        read_def_rpr_into(e, &mut style);
                        in_def_rpr = true;
                    }
                    b"solidFill" if in_def_rpr => in_solid_fill = true,
                    // A self-closing colour element has no `<a:alpha>` child,
                    // so this arm leaves the opacity alone (issue #1677).
                    _ if in_solid_fill && style.color.is_none() => {
                        style.color = drawingml::parse_color_from_empty(e, scheme).color;
                    }
                    _ => {}
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"txPr" => break,
                b"defRPr" => in_def_rpr = false,
                b"solidFill" => in_solid_fill = false,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }
    (typeface, style, ellipsis)
}

/// Read a `c:txPr` that governs one element rather than the whole chart space,
/// keeping only its run properties.
fn parse_chart_text_style(reader: &mut Reader<&[u8]>, scheme: &SchemeColors<'_>) -> ChartTextStyle {
    parse_chart_text_properties(reader, scheme).1
}

/// Read an axis typeface, run properties, and body overflow policy.
fn parse_axis_text_properties(
    reader: &mut Reader<&[u8]>,
    scheme: &SchemeColors<'_>,
) -> (Option<String>, ChartTextStyle) {
    let (typeface, mut style, ellipsis) = parse_chart_text_properties(reader, scheme);
    style.ellipsis_overflow = ellipsis;
    (typeface, style)
}

/// Take `a:defRPr`'s run properties, leaving untouched whatever it omits.
///
/// Only the first `a:defRPr` counts: `c:txPr` holds one paragraph, and a later
/// `a:endParaRPr` describes the empty run after the text rather than the text.
fn read_def_rpr_into(element: &quick_xml::events::BytesStart<'_>, style: &mut ChartTextStyle) {
    if style.size_pt.is_none() {
        // `@sz` is in hundredths of a point.
        style.size_pt = xml_util::get_attr_str(element, b"sz")
            .and_then(|raw| raw.parse::<f64>().ok())
            .filter(|hundredths| *hundredths > 0.0)
            .map(|hundredths| hundredths / 100.0);
    }
    if style.bold.is_none() {
        style.bold =
            xml_util::get_attr_str(element, b"b").map(|raw| matches!(raw.trim(), "1" | "true"));
    }
    if style.letter_spacing_hundredths.is_none() {
        // DrawingML character spacing is in hundredths of a point and may be
        // negative. Unlike an absent attribute, an explicit zero is still a
        // declaration that overrides the chart-space default.
        style.letter_spacing_hundredths =
            xml_util::get_attr_str(element, b"spc").and_then(|raw| raw.parse::<i32>().ok());
    }
}

/// Read `c:chartSpace/c:spPr` into the chart area's fill and outline.
///
/// A top-level fill paints the whole chart, while a fill nested in `a:ln`
/// colours only the outline. Keeping those contexts separate prevents the
/// line colour from becoming the background when no area fill exists (#1217).
/// The outline's absent / no-fill / explicit states remain as defined in #637.
fn parse_chart_area_properties(
    reader: &mut Reader<&[u8]>,
    scheme: &SchemeColors<'_>,
) -> (ChartAreaFill, ChartAreaOutline) {
    let mut fill: ChartAreaFill = ChartAreaFill::Unspecified;
    let mut in_line: bool = false;
    let mut saw_line: bool = false;
    let mut suppressed: bool = false;
    let mut width_pt: Option<f64> = None;
    let mut line_color: Option<Color> = None;
    let mut in_area_solid_fill: bool = false;
    let mut in_line_solid_fill: bool = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"ln" =>
            {
                in_line = true;
                saw_line = true;
                width_pt = width_pt.or_else(|| {
                    xml_util::get_attr_str(e, b"w")
                        .and_then(|w| w.parse::<f64>().ok())
                        .map(|emu| emu / EMU_PER_POINT)
                });
            }
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if in_line && e.local_name().as_ref() == b"noFill" =>
            {
                suppressed = true;
            }
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if !in_line && e.local_name().as_ref() == b"noFill" =>
            {
                fill = ChartAreaFill::Transparent;
            }
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                if in_line {
                    in_line_solid_fill = true;
                } else {
                    in_area_solid_fill = true;
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                in_area_solid_fill = false;
                in_line_solid_fill = false;
            }
            Ok(Event::Start(ref e))
                if (in_area_solid_fill || in_line_solid_fill)
                    && matches!(e.local_name().as_ref(), b"srgbClr" | b"schemeClr") =>
            {
                let parsed = drawingml::parse_color_from_start(reader, e, scheme).color;
                if in_line_solid_fill {
                    line_color = line_color.or(parsed);
                } else if let Some(color) = parsed {
                    fill = ChartAreaFill::Solid(color);
                }
            }
            Ok(Event::Empty(ref e))
                if (in_area_solid_fill || in_line_solid_fill)
                    && matches!(e.local_name().as_ref(), b"srgbClr" | b"schemeClr") =>
            {
                let parsed = drawingml::parse_color_from_empty(e, scheme).color;
                if in_line_solid_fill {
                    line_color = line_color.or(parsed);
                } else if let Some(color) = parsed {
                    fill = ChartAreaFill::Solid(color);
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"ln" => {
                in_line = false;
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"spPr" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    let outline: ChartAreaOutline = if !saw_line {
        ChartAreaOutline::Default
    } else if suppressed {
        ChartAreaOutline::Suppressed
    } else {
        ChartAreaOutline::Explicit {
            width_pt,
            color: line_color,
        }
    };

    (fill, outline)
}

/// EMU in one point. `a:ln/@w` is in EMU.
const EMU_PER_POINT: f64 = 12700.0;

/// The axes a plot area's families read against, settled by reference.
struct AxisGroups {
    /// The category axis of the primary group — the one whose furniture is
    /// drawn.
    category_axis: Axis,
    /// The value axis of the primary group.
    value_axis: Axis,
    /// The value axis of a second group, when some family reads against it.
    secondary_value_axis: Option<Axis>,
}

/// Settle which `<c:catAx>` and `<c:valAx>` each family plots against.
///
/// The primary group is the one the first family references through its
/// `<c:axId>` pair — Office writes the primary group's families first — and
/// falls back to the first axis of each kind for a part naming no ids, which
/// is what every single-group chart amounts to. A family referencing another
/// `<c:valAx>` puts its series on the secondary axis, which is kept only when
/// something reads against it. Document order of the axis elements plays no
/// part: the #1220 deck writes its axes primary-first, but nothing requires
/// that (issue #1374).
///
/// TODO(three value axes): ECMA-376 allows only two axis groups in a 2-D plot
/// area, so a third referenced `<c:valAx>` is read as the primary one.
fn resolve_axis_groups(
    mut category_axes: Vec<Axis>,
    mut value_axes: Vec<Axis>,
    family_axis_ids: &[(std::ops::Range<usize>, Vec<String>)],
    series: &mut [ChartSeries],
) -> AxisGroups {
    fn referenced<'a>(ids: &'a [String], axes: &[Axis]) -> Option<&'a str> {
        ids.iter()
            .map(String::as_str)
            .find(|id| axes.iter().any(|axis| axis.id.as_deref() == Some(*id)))
    }
    fn take(axes: &mut Vec<Axis>, id: Option<&str>) -> Option<Axis> {
        let index: usize = axes
            .iter()
            .position(|axis| axis.id.is_some() && axis.id.as_deref() == id)?;
        Some(axes.remove(index))
    }
    fn take_or_first(axes: &mut Vec<Axis>, id: Option<&str>) -> Axis {
        take(axes, id)
            .or_else(|| (!axes.is_empty()).then(|| axes.remove(0)))
            .unwrap_or_default()
    }

    let primary_ids: &[String] = family_axis_ids
        .first()
        .map(|(_, ids)| ids.as_slice())
        .unwrap_or(&[]);
    let primary_value_id: Option<String> = referenced(primary_ids, &value_axes)
        .map(str::to_string)
        .or_else(|| value_axes.first().and_then(|axis| axis.id.clone()));
    let primary_category_id: Option<String> = referenced(primary_ids, &category_axes)
        .map(str::to_string)
        .or_else(|| category_axes.first().and_then(|axis| axis.id.clone()));

    let mut secondary_value_id: Option<String> = None;
    for (range, ids) in family_axis_ids {
        let Some(id) = referenced(ids, &value_axes) else {
            continue;
        };
        if primary_value_id.as_deref() == Some(id) {
            continue;
        }
        let secondary_id: &str = secondary_value_id.get_or_insert_with(|| id.to_string());
        if secondary_id != id {
            continue;
        }
        for entry in &mut series[range.clone()] {
            entry.value_axis = ChartValueAxisRole::Secondary;
        }
    }

    let secondary_value_axis: Option<Axis> = take(&mut value_axes, secondary_value_id.as_deref());
    AxisGroups {
        category_axis: take_or_first(&mut category_axes, primary_category_id.as_deref()),
        value_axis: take_or_first(&mut value_axes, primary_value_id.as_deref()),
        secondary_value_axis,
    }
}

/// What one `<c:catAx>` or `<c:valAx>` element says about itself.
#[derive(Default)]
struct Axis {
    /// `<c:axId>` — the identity a chart family references this axis by.
    id: Option<String>,
    /// `<c:axPos>` — the plot edge the axis runs along: `l`, `r`, `t` or `b`.
    position: Option<String>,
    font_family: Option<String>,
    title: Option<String>,
    major_tick_mark: AxisTickMark,
    deleted: bool,
    text_style: ChartTextStyle,
    /// `<c:numFmt formatCode>` — how this axis prints its tick labels.
    number_format: Option<String>,
    /// What `<c:spPr>` says about the axis' own line.
    line: ChartLine,
    /// Presence and stroke of `<c:majorGridlines>`. `None` means absent;
    /// `Some(Automatic)` means present without a declared stroke.
    gridline: Option<ChartLine>,
    /// `<c:majorUnit>` — the tick interval this axis states.
    major_unit: Option<f64>,
    /// `<c:scaling><c:min>` — the value this axis starts at.
    min: Option<f64>,
    /// `<c:scaling><c:max>` — the value this axis ends at.
    max: Option<f64>,
}

/// The `formatCode` a `<c:numFmt>` states, when it states one that is not
/// `General`.
///
/// `General` is Excel's "no format" and applying it would only reformat the
/// number the plain path already prints (issue #865).
fn explicit_format_code(element: &quick_xml::events::BytesStart<'_>) -> Option<String> {
    let code = xml_util::get_attr_str(element, b"formatCode")?;
    let code = code.trim();
    (!code.is_empty() && !code.eq_ignore_ascii_case("General")).then(|| code.to_string())
}

/// Read a `<c:scaling>` body, reporting the interval it fixes the axis to.
///
/// `<c:min>` and `<c:max>` are optional and independent: a part may fix one
/// end and leave the other to the automatic scale. A bound that is not a
/// finite number is dropped rather than propagated as a NaN that would make
/// every plotted position vanish (issue #1184).
///
/// TODO(orientation): `<c:orientation val="maxMin"/>` reverses the axis, which
/// the renderer does not model, so it is read past here.
fn parse_axis_scaling(reader: &mut Reader<&[u8]>) -> (Option<f64>, Option<f64>) {
    let mut min: Option<f64> = None;
    let mut max: Option<f64> = None;
    let bound = |element: &quick_xml::events::BytesStart<'_>| -> Option<f64> {
        xml_util::get_attr_str(element, b"val")
            .and_then(|value| value.parse::<f64>().ok())
            .filter(|value| value.is_finite())
    };
    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"min" => min = bound(e).or(min),
                b"max" => max = bound(e).or(max),
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"scaling" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    (min, max)
}

/// Read an axis element, consuming it to `end_tag`.
///
/// Axis titles sit after the plot-area family element, so they never reach the
/// `<c:title>` branch that captures the chart's own title.
fn parse_axis(reader: &mut Reader<&[u8]>, end_tag: &[u8], scheme: &SchemeColors<'_>) -> Axis {
    let mut axis: Axis = Axis::default();
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"title" => {
                axis.title = axis.title.or_else(|| parse_chart_title(reader, scheme).0);
            }
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"txPr" => {
                (axis.font_family, axis.text_style) = parse_axis_text_properties(reader, scheme);
            }
            // The axis' own line, and the gridlines' — both are an `<a:ln>`
            // inside a `<c:spPr>`, and both are dropped without this (#900).
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                axis.line = parse_chart_line(reader, b"spPr", scheme);
            }
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"majorGridlines" => {
                axis.gridline = Some(parse_chart_line(reader, b"majorGridlines", scheme));
            }
            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"majorGridlines" => {
                axis.gridline = Some(ChartLine::Automatic);
            }
            // `<c:scaling>` is the only place an axis states a bound, and it
            // holds `<c:orientation>` too, so it is consumed as a body rather
            // than matched flat: `min`/`max` mean nothing outside it (#1184).
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"scaling" => {
                let (min, max) = parse_axis_scaling(reader);
                axis.min = min;
                axis.max = max;
            }
            // Office writes `<c:majorTickMark val="out"/>` self-closing, so the
            // `Start` arm alone would never see it.
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.local_name().as_ref() == b"majorTickMark" =>
            {
                axis.major_tick_mark = xml_util::get_attr_str(e, b"val")
                    .as_deref()
                    .map(axis_tick_mark_for)
                    .unwrap_or_default();
            }
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.local_name().as_ref() == b"majorUnit" =>
            {
                axis.major_unit = xml_util::get_attr_str(e, b"val")
                    .and_then(|value| value.parse::<f64>().ok())
                    .filter(|unit| unit.is_finite() && *unit > 0.0);
            }
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.local_name().as_ref() == b"delete" =>
            {
                axis.deleted = ct_boolean(e);
            }
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.local_name().as_ref() == b"axId" =>
            {
                axis.id = xml_util::get_attr_str(e, b"val");
            }
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.local_name().as_ref() == b"axPos" =>
            {
                axis.position = xml_util::get_attr_str(e, b"val");
            }
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                if e.local_name().as_ref() == b"numFmt" =>
            {
                axis.number_format = explicit_format_code(e);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == end_tag => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    axis
}

/// Consume a `<c:legend>` body, reporting whether it switches itself off and
/// which edge it names, and the run properties its entries take.
///
/// `<c:legend><c:delete val="1"/></c:legend>` is how Office records a legend
/// that was turned off but whose settings were kept, the same shape the axes
/// use. The position comes back from here too because this consumes the body,
/// so `<c:legendPos>` never reaches the caller's loop (issue #762).
fn parse_legend(
    reader: &mut Reader<&[u8]>,
    scheme: &SchemeColors<'_>,
) -> (bool, Option<LegendPosition>, ChartTextStyle) {
    let mut deleted = false;
    let mut position: Option<LegendPosition> = None;
    let mut style: ChartTextStyle = ChartTextStyle::default();
    loop {
        match reader.read_event() {
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"delete" => deleted = ct_boolean(e),
                b"legendPos" => {
                    position = xml_util::get_attr_str(e, b"val")
                        .as_deref()
                        .map(legend_position_for);
                }
                _ => {}
            },
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                // A `c:legendEntry` can carry a `c:txPr` for one overridden
                // entry. Do not promote that narrower scope to every entry in
                // the legend merely because we encounter it first.
                b"legendEntry" => xml_util::skip_element(reader, b"legendEntry"),
                b"txPr" => style = parse_chart_text_style(reader, scheme),
                _ => {}
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"legend" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    (deleted, position, style)
}

/// Read a `CT_Boolean` element's own state.
///
/// ECMA-376 defaults `val` to true, so a bare `<c:delete/>` or `<c:showVal/>`
/// turns the flag on.
fn ct_boolean(element: &quick_xml::events::BytesStart) -> bool {
    xml_util::get_attr_str(element, b"val")
        .map(|value| value == "1" || value == "true")
        .unwrap_or(true)
}

/// Resolve a `<c:majorTickMark>` value to the side its ticks reach from.
fn axis_tick_mark_for(value: &str) -> AxisTickMark {
    match value {
        "none" => AxisTickMark::None,
        "in" => AxisTickMark::Inside,
        "cross" => AxisTickMark::Cross,
        _ => AxisTickMark::Outside,
    }
}

/// Which run properties the reader is inside while reading a `<c:rich>`.
///
/// `<a:pPr><a:defRPr>` states the paragraph's defaults and `<a:r><a:rPr>` one
/// run's overrides; both carry an `<a:solidFill>`, so the colour has to know
/// which of the two it belongs to.
#[derive(Clone, Copy, PartialEq, Eq)]
enum RichRunScope {
    Paragraph,
    Run,
}

/// Read a title's `<c:rich>`: the text it names and the run properties that
/// paint it.
///
/// PowerPoint writes a chart title's formatting here, not in the `<c:txPr>`
/// beside it: `chart8.xml` of the deck in issue #1220 states `b="1"` over
/// `<a:schemeClr val="accent2">` on the run, while its `<c:txPr>` repeats the
/// regular grey paragraph default for the empty run after the text. The
/// subtree was read for its text alone, so both titles printed regular grey.
///
/// A title is one string in the model, so where several runs state the same
/// property the first one that states it wins — the rule `<c:txPr>` already
/// follows for a repeated `<a:defRPr>`.
fn parse_chart_rich_text(
    reader: &mut Reader<&[u8]>,
    scheme: &SchemeColors<'_>,
) -> (String, ChartTextStyle) {
    let mut text = String::new();
    let mut in_t: bool = false;
    let mut paragraph: ChartTextStyle = ChartTextStyle::default();
    let mut run: ChartTextStyle = ChartTextStyle::default();
    let mut scope: Option<RichRunScope> = None;
    let mut in_p_pr: bool = false;
    let mut in_run: bool = false;
    let mut in_solid_fill: bool = false;
    // `<a:ln>` inside the run properties outlines the glyphs, and its own
    // `<a:solidFill>` is not the colour the text is painted in (issue #916).
    let mut in_line: bool = false;

    loop {
        match reader.read_event() {
            // A colour element written with children keeps its transforms
            // there, so its start tag alone is not the colour; this arm has to
            // precede the shared start/empty one below, which cannot consume
            // the children (issue #1160).
            Ok(Event::Start(ref e))
                if in_solid_fill
                    && matches!(
                        e.local_name().as_ref(),
                        b"srgbClr" | b"schemeClr" | b"sysClr"
                    ) =>
            {
                let parsed = drawingml::parse_color_from_start(reader, e, scheme);
                if let Some(target) = rich_scope_style(scope, &mut paragraph, &mut run) {
                    // The colour and its `<a:alpha>` child move together
                    // (issue #1677).
                    if target.color.is_none() && parsed.color.is_some() {
                        target.color = parsed.color;
                        target.alpha = parsed.alpha;
                    }
                }
            }
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"pPr" => in_p_pr = true,
                b"r" => in_run = true,
                b"ln" => in_line = true,
                b"defRPr" if in_p_pr => {
                    read_def_rpr_into(e, &mut paragraph);
                    scope = Some(RichRunScope::Paragraph);
                }
                b"rPr" if in_run => {
                    read_def_rpr_into(e, &mut run);
                    scope = Some(RichRunScope::Run);
                }
                b"solidFill" if scope.is_some() && !in_line => in_solid_fill = true,
                b"t" => in_t = true,
                _ => {}
            },
            Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                // A self-closing property element opens no scope: `<a:ln>`'s
                // fill would otherwise still be read as the text colour.
                b"defRPr" if in_p_pr => read_def_rpr_into(e, &mut paragraph),
                b"rPr" if in_run => read_def_rpr_into(e, &mut run),
                _ if in_solid_fill => {
                    if let Some(target) = rich_scope_style(scope, &mut paragraph, &mut run) {
                        target.color = target
                            .color
                            .or(drawingml::parse_color_from_empty(e, scheme).color);
                    }
                }
                _ => {}
            },
            Ok(Event::Text(ref t)) if in_t => {
                if let Ok(s) = t.xml_content(OOXML_XML_VERSION) {
                    text.push_str(s.as_ref());
                }
            }
            Ok(Event::GeneralRef(ref reference)) if in_t => {
                if let Some(s) = xml_util::decode_general_ref(reference) {
                    text.push_str(&s);
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"rich" => break,
                b"pPr" => in_p_pr = false,
                // `<a:endParaRPr>` describes the empty run after the text, so
                // leaving the run here keeps it out of the title's style.
                b"r" => in_run = false,
                b"ln" => in_line = false,
                b"defRPr" | b"rPr" => scope = None,
                b"solidFill" => in_solid_fill = false,
                b"t" => in_t = false,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    (text, paragraph.overridden_by(run))
}

/// The style the run properties currently being read belong to.
fn rich_scope_style<'a>(
    scope: Option<RichRunScope>,
    paragraph: &'a mut ChartTextStyle,
    run: &'a mut ChartTextStyle,
) -> Option<&'a mut ChartTextStyle> {
    match scope {
        Some(RichRunScope::Paragraph) => Some(paragraph),
        Some(RichRunScope::Run) => Some(run),
        None => None,
    }
}

/// Parse the chart title text and its own run properties from `<c:title>`.
///
/// Returns the text, whether the element named text of its own — a `<c:tx>` —,
/// what its `<c:txPr>` declares, and any deterministic manual edge anchor. A
/// title without a `<c:tx>` is *automatic*: it carries the formatting for a
/// string the application supplies (issue #1146), so an empty result there
/// means something different from an empty `<c:tx>`.
///
/// The `<c:txPr>` is the title's own, and outranks the chart space's for the
/// string it governs: `any_sheets.xlsx` states `sz="1400" b="0"` in grey there
/// beside a chart space that states nothing at all (issue #1215). The
/// `<c:tx><c:rich>` runs that carry the string are more specific still, and
/// outrank that `<c:txPr>` in turn (issue #1424).
fn parse_chart_title(
    reader: &mut Reader<&[u8]>,
    scheme: &SchemeColors<'_>,
) -> (
    Option<String>,
    bool,
    ChartTextStyle,
    Option<ChartTitleLayout>,
) {
    let mut text = String::new();
    let mut in_t = false;
    let mut names_own_text = false;
    let mut tx_pr_style: ChartTextStyle = ChartTextStyle::default();
    let mut rich_style: ChartTextStyle = ChartTextStyle::default();
    let mut layout: Option<ChartTitleLayout> = None;
    let mut depth = 1u32;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                let local = e.local_name();
                if local.as_ref() == b"title" {
                    depth += 1;
                } else if local.as_ref() == b"txPr" {
                    // Consumes through `</c:txPr>`, so the reader comes back on
                    // the title's next sibling.
                    tx_pr_style = parse_chart_text_style(reader, scheme);
                } else if local.as_ref() == b"layout" {
                    layout = parse_title_layout(reader);
                } else if local.as_ref() == b"tx" {
                    names_own_text = true;
                } else if local.as_ref() == b"rich" {
                    // Consumes through `</c:rich>`, text included, so `a:t`
                    // never reaches the arm below.
                    let (rich_text, style) = parse_chart_rich_text(reader, scheme);
                    text.push_str(&rich_text);
                    rich_style = rich_style.overridden_by(style);
                } else if local.as_ref() == b"t" {
                    in_t = true;
                }
            }
            Ok(Event::Text(ref t)) if in_t => {
                if let Ok(s) = t.xml_content(OOXML_XML_VERSION) {
                    text.push_str(s.as_ref());
                }
            }
            Ok(Event::GeneralRef(ref reference)) if in_t => {
                if let Some(s) = xml_util::decode_general_ref(reference) {
                    text.push_str(&s);
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
    let title: Option<String> = if trimmed.is_empty() {
        None
    } else {
        Some(trimmed)
    };
    (
        title,
        names_own_text,
        tx_pr_style.overridden_by(rich_style),
        layout,
    )
}

/// Plot-area settings that sit beside `<c:ser>` inside a chart type element.
///
/// Office writes them all self-closing, so they arrive as `Empty` events.
#[derive(Debug, Default)]
struct PlotAreaProps {
    /// `<c:barDir>`, exclusive to the bar family.
    bar_direction: Option<String>,
    /// `<c:grouping>`.
    grouping: Option<String>,
    /// `<c:radarStyle>` selects filled versus line-based radar styling.
    radar_style: Option<String>,
    /// `<c:ofPieType>`, exclusive to `<c:ofPieChart>`.
    of_pie_type: Option<String>,
    /// `<c:gapWidth>`, exclusive to the bar family. Office writes it after the
    /// last `</c:ser>`, so it lands here rather than on a series.
    gap_width: Option<String>,
    /// `<c:overlap>`, exclusive to the bar family, and likewise trailing.
    overlap: Option<String>,
    /// `<c:holeSize>`, exclusive to the doughnut family (issue #679).
    hole_size: Option<String>,
    /// `<c:firstSliceAng>`, declared by the pie and doughnut families alone,
    /// and written after the last `</c:ser>` as `<c:gapWidth>` is
    /// (issue #1429).
    first_slice_ang: Option<String>,
    /// `<c:axId>`, one per axis the family plots against, trailing the series
    /// like `<c:gapWidth>` does (issue #1374).
    axis_ids: Vec<String>,
}

impl PlotAreaProps {
    /// Record `e` when it is one of the settings this struct carries.
    fn absorb(&mut self, e: &quick_xml::events::BytesStart) -> bool {
        match e.local_name().as_ref() {
            b"barDir" => self.bar_direction = xml_util::get_attr_str(e, b"val"),
            b"grouping" => self.grouping = xml_util::get_attr_str(e, b"val"),
            b"radarStyle" => self.radar_style = xml_util::get_attr_str(e, b"val"),
            b"ofPieType" => self.of_pie_type = xml_util::get_attr_str(e, b"val"),
            b"gapWidth" => self.gap_width = xml_util::get_attr_str(e, b"val"),
            b"overlap" => self.overlap = xml_util::get_attr_str(e, b"val"),
            b"holeSize" => self.hole_size = xml_util::get_attr_str(e, b"val"),
            b"firstSliceAng" => self.first_slice_ang = xml_util::get_attr_str(e, b"val"),
            b"axId" => self.axis_ids.extend(xml_util::get_attr_str(e, b"val")),
            _ => return false,
        }
        true
    }

    /// The band layout this element declares, as far as it declares one.
    ///
    /// Nothing is substituted for an absent or unreadable element: the caller
    /// keeps looking, and only a chart that declared nothing anywhere falls
    /// back to [`BarBandLayout::default`].
    fn bar_band_layout(&self) -> (Option<f64>, Option<f64>) {
        (
            self.gap_width
                .as_deref()
                .and_then(|value| bar_percent(value, 0.0, 500.0)),
            self.overlap
                .as_deref()
                .and_then(|value| bar_percent(value, -100.0, 100.0)),
        )
    }
}

/// Read `<c:gapWidth>`'s or `<c:overlap>`'s `val`, held to the range its type
/// allows.
///
/// `ST_GapAmount` and `ST_Overlap` are each a union of a bare integer and a
/// percentage string, so `"90"` and `"90%"` describe the same chart. Office
/// enforces the ranges itself — PowerPoint 16.0 refuses to open a file whose
/// gapWidth reads 1000 while opening 500 happily — so a value outside them
/// describes no drawable chart and the nearest bound is the closest reading of
/// what it meant.
fn bar_percent(value: &str, low: f64, high: f64) -> Option<f64> {
    value
        .trim()
        .trim_end_matches('%')
        .trim()
        .parse::<f64>()
        .ok()
        .filter(|percent| percent.is_finite())
        .map(|percent| percent.clamp(low, high))
}

/// Read `<c:firstSliceAng>`'s `val` as degrees clockwise of twelve o'clock.
///
/// `ST_FirstSliceAng` is an unsigned integer the schema bounds to 0..360, and
/// PowerPoint's own dial offers exactly that range. A file outside it
/// describes no drawable rotation, so the nearest bound is the closest reading
/// of what it meant — the same treatment `<c:gapWidth>` gets in
/// [`bar_percent`] (issue #1429).
fn first_slice_angle_degrees(value: &str) -> Option<u32> {
    value
        .trim()
        .parse::<i64>()
        .ok()
        .map(|degrees| degrees.clamp(0, 360) as u32)
}

/// Parse series data from within a chart type element (e.g., `<c:barChart>`).
fn parse_chart_series(
    reader: &mut Reader<&[u8]>,
    end_tag: &[u8],
    categories: &mut Vec<String>,
    series: &mut Vec<ChartSeries>,
    plot: &mut PlotAreaProps,
    scheme: &SchemeColors<'_>,
) {
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"ser" {
                    let (ser, cats) = parse_single_series(reader, scheme);
                    // Use categories from first series that has them
                    if categories.is_empty() && !cats.is_empty() {
                        *categories = cats;
                    }
                    series.push(ser);
                } else {
                    plot.absorb(e);
                }
            }
            Ok(Event::Empty(ref e)) => {
                plot.absorb(e);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == end_tag => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

/// Parse a single `<c:ser>` element and return the series data + category labels.
fn parse_single_series(
    reader: &mut Reader<&[u8]>,
    scheme: &SchemeColors<'_>,
) -> (ChartSeries, Vec<String>) {
    let mut name = None;
    let mut values = Vec::new();
    let mut categories = Vec::new();
    let mut fill: Option<Color> = None;
    let mut fill_mode = ChartFillMode::Automatic;
    let mut data_labels: DataLabels = DataLabels::default();
    // Keyed by `<c:idx>` because `<c:dPt>` entries are written only for the
    // points that override the series, and not necessarily in order.
    let mut point_fills: std::collections::BTreeMap<usize, (Option<Color>, ChartFillMode)> =
        std::collections::BTreeMap::new();
    let mut number_format: Option<String> = None;
    let mut marker_symbol: Option<MarkerSymbol> = None;
    let mut marker_style = ChartMarkerStyle::default();
    let mut line_width_pt: Option<f64> = None;
    let mut line_geometry = crate::ir::ChartStrokeGeometry::default();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"tx" => name = parse_series_text(reader),
                b"cat" => categories = parse_category_data(reader),
                b"val" | b"yVal" => {
                    let (parsed, format_code) = parse_value_data(reader);
                    values = parsed;
                    number_format = format_code;
                }
                // The series' own fill and stroke geometry. This match is flat,
                // so it would also see a `<c:spPr>` nested inside a sibling
                // element; every element that can carry one is consumed by its
                // own branch first, leaving only the series-level one here.
                b"spPr" => {
                    let properties = parse_shape_properties(reader, b"spPr", scheme);
                    fill = fill.or(properties.fill);
                    if fill_mode == ChartFillMode::Automatic {
                        fill_mode = properties.fill_mode;
                    }
                    line_width_pt = line_width_pt.or(properties.line_width_pt);
                    line_geometry.cap = line_geometry.cap.or(properties.line_geometry.cap);
                    line_geometry.join = line_geometry.join.or(properties.line_geometry.join);
                    line_geometry.miter_limit = line_geometry
                        .miter_limit
                        .or(properties.line_geometry.miter_limit);
                }
                b"dPt" => {
                    if let Some((index, color, mode)) = parse_data_point(reader, scheme) {
                        point_fills.insert(index, (color, mode));
                    }
                }
                // Consumed whole for the same reason as `<c:dLbls>`: the
                // marker carries an `<c:spPr>` for the symbol's own fill.
                b"marker" => (marker_symbol, marker_style) = parse_series_marker(reader, scheme),
                // Consumed whole: `<c:dLbls>` carries an `<c:spPr>` of its
                // own for the label box, which would otherwise be read as the
                // series fill.
                b"dLbls" => data_labels = parse_data_labels(reader, scheme),
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

    let point_count: usize = point_fills.keys().last().map_or(0, |last| last + 1);
    let point_fill_modes: Vec<ChartFillMode> = (0..point_count)
        .map(|index| {
            point_fills
                .get(&index)
                .map_or(ChartFillMode::Automatic, |entry| entry.1)
        })
        .collect();
    let point_fills: Vec<Option<Color>> = (0..point_count)
        .map(|index| point_fills.get(&index).and_then(|entry| entry.0))
        .collect();

    (
        ChartSeries {
            name,
            values,
            fill,
            fill_mode,
            point_fills,
            point_fill_modes,
            data_labels,
            number_format,
            // Filled in by the caller, which knows the family element this
            // series was read inside (issue #1067).
            plot_type: None,
            value_axis: ChartValueAxisRole::Primary,
            marker_symbol,
            marker_style,
            line_width_pt,
            line_geometry,
        },
        categories,
    )
}

/// Read a `<c:dLbls>` element, consuming it whole.
///
/// `CT_DLbls` is `dLbl*` followed by the group-level settings, and a per-point
/// `<c:dLbl>` repeats the same `showVal`/`showCatName`/… names. This loop
/// matches on local name alone, so each `<c:dLbl>` is skipped whole; reading
/// them would let a single point's override become the series default whenever
/// the group-level settings are absent.
fn parse_data_labels(reader: &mut Reader<&[u8]>, scheme: &SchemeColors<'_>) -> DataLabels {
    let mut labels = DataLabels::default();
    let mut in_separator: bool = false;
    let mut separator = String::new();
    // `<c:dLblPos>` is optional; `None` here means the grouping's default is
    // applied later, once the plot family is known (issue #901).
    let mut stated_position: Option<DataLabelPosition> = None;

    loop {
        let event = reader.read_event();
        match event {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => match e.local_name().as_ref() {
                b"dLbl" => xml_util::skip_element(reader, b"dLbl"),
                b"showVal" => labels.show_value = ct_boolean(e),
                b"showCatName" => labels.show_category = ct_boolean(e),
                b"showSerName" => labels.show_series = ct_boolean(e),
                b"showPercent" => labels.show_percent = ct_boolean(e),
                b"numFmt" => labels.number_format = explicit_format_code(e),
                b"dLblPos" => {
                    stated_position = xml_util::get_attr_str(e, b"val")
                        .as_deref()
                        .and_then(data_label_position_for);
                }
                b"separator" => in_separator = true,
                // Only the group-level `c:txPr` reaches here — a per-point
                // `<c:dLbl>` carrying one of its own is skipped whole above.
                b"txPr" => labels.text_style = parse_chart_text_style(reader, scheme),
                _ => {}
            },
            Ok(Event::Text(ref text)) if in_separator => {
                if let Ok(value) = text.xml_content(OOXML_XML_VERSION) {
                    separator.push_str(value.as_ref());
                }
            }
            Ok(Event::GeneralRef(ref reference)) if in_separator => {
                if let Some(value) = xml_util::decode_general_ref(reference) {
                    separator.push_str(&value);
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"separator" => in_separator = false,
                b"dLbls" => break,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if !separator.is_empty() {
        labels.separator = separator;
    }
    // A position the part did not state is settled by the grouping, which the
    // caller knows; `Center` stands in until then and matches the stacked
    // default, so only a clustered plot has to override it.
    if let Some(position) = stated_position {
        labels.position = position;
        labels.position_stated = true;
    }
    labels
}

/// `<c:dLblPos val>` as ECMA-376 §21.2.2.49 names the positions. `bestFit`
/// and the pie-only `t`/`b`/`l`/`r` are not mapped: they describe placements
/// this renderer does not draw, and guessing one would move a label the file
/// did not ask to move.
fn data_label_position_for(value: &str) -> Option<DataLabelPosition> {
    match value {
        "ctr" => Some(DataLabelPosition::Center),
        "outEnd" => Some(DataLabelPosition::OutsideEnd),
        "inEnd" => Some(DataLabelPosition::InsideEnd),
        "inBase" => Some(DataLabelPosition::InsideBase),
        _ => None,
    }
}

/// Read a series marker's symbol, size, and independent shape styling.
///
/// Consumed whole whether or not it names one this renderer draws: the element
/// carries an `<c:spPr>` for the symbol's own fill, which the flat series loop
/// would otherwise read as the series fill.
///
/// The `<c:marker val="1"/>` that sits beside the `<c:ser>` elements of a
/// `<c:lineChart>` is a different element — `CT_Boolean`, saying whether the
/// family shows markers at all — and never reaches here, being empty.
fn parse_series_marker(
    reader: &mut Reader<&[u8]>,
    scheme: &SchemeColors<'_>,
) -> (Option<MarkerSymbol>, ChartMarkerStyle) {
    let mut symbol: Option<MarkerSymbol> = None;
    let mut style = ChartMarkerStyle::default();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"symbol" =>
            {
                symbol = xml_util::get_attr_str(e, b"val")
                    .as_deref()
                    .and_then(marker_symbol_for);
            }
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"size" =>
            {
                style.size_pt = xml_util::get_attr_str(e, b"val")
                    .and_then(|v| v.parse::<u8>().ok())
                    .filter(|v| (2..=72).contains(v))
                    .map(f64::from);
            }
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                parse_marker_properties(reader, scheme, &mut style);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"marker" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    (symbol, style)
}

/// Marker fill and outline are separate paints; neither may leak into the
/// series colour or into the other paint when the first is absent.
fn parse_marker_properties(
    reader: &mut Reader<&[u8]>,
    scheme: &SchemeColors<'_>,
    style: &mut ChartMarkerStyle,
) {
    let mut in_line = false;
    let mut in_solid_fill = false;
    loop {
        let event = reader.read_event();
        let starts_line = matches!(&event, Ok(Event::Start(e)) if e.local_name().as_ref() == b"ln");
        match event {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"ln" =>
            {
                in_line = starts_line;
                style.line = ChartLine::Explicit {
                    width_pt: xml_util::get_attr_str(e, b"w")
                        .and_then(|v| v.parse::<f64>().ok())
                        .filter(|v| v.is_finite() && *v >= 0.0)
                        .map(|v| v / EMU_PER_POINT),
                    color: None,
                    alpha: None,
                };
            }
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"noFill" =>
            {
                if in_line {
                    style.line = ChartLine::Suppressed;
                } else {
                    style.fill_mode = ChartFillMode::Suppressed;
                }
            }
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                in_solid_fill = true;
                if !in_line {
                    style.fill_mode = ChartFillMode::Explicit;
                }
            }
            Ok(Event::Start(ref e))
                if in_solid_fill
                    && matches!(e.local_name().as_ref(), b"srgbClr" | b"schemeClr") =>
            {
                let parsed = drawingml::parse_color_from_start(reader, e, scheme);
                set_marker_paint(style, in_line, parsed.color, parsed.alpha);
            }
            Ok(Event::Empty(ref e))
                if in_solid_fill
                    && matches!(e.local_name().as_ref(), b"srgbClr" | b"schemeClr") =>
            {
                let parsed = drawingml::parse_color_from_empty(e, scheme);
                set_marker_paint(style, in_line, parsed.color, parsed.alpha);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                in_solid_fill = false
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"ln" => in_line = false,
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"spPr" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
}

fn set_marker_paint(
    style: &mut ChartMarkerStyle,
    in_line: bool,
    color: Option<Color>,
    alpha: Option<f64>,
) {
    if in_line {
        if let ChartLine::Explicit {
            color: line_color,
            alpha: line_alpha,
            ..
        } = &mut style.line
        {
            *line_color = color;
            *line_alpha = alpha;
        }
    } else {
        style.fill = color;
        style.fill_alpha = alpha;
    }
}

/// `<c:symbol val>` as ECMA-376 §21.2.3.29 `ST_MarkerStyle` names the symbols.
///
/// `auto` is the file asking for the automatic symbol, so it maps to `None`
/// like an absent element. `dash`, `dot`, `plus`, `star` and `picture` map
/// there too, for want of a shape to draw them as: the automatic cycle at
/// least keeps adjacent series apart, where substituting one named symbol for
/// another would state something the file did not.
fn marker_symbol_for(value: &str) -> Option<MarkerSymbol> {
    match value {
        "none" => Some(MarkerSymbol::Off),
        "circle" => Some(MarkerSymbol::Circle),
        "diamond" => Some(MarkerSymbol::Diamond),
        "square" => Some(MarkerSymbol::Square),
        "triangle" => Some(MarkerSymbol::Triangle),
        "x" => Some(MarkerSymbol::Cross),
        _ => None,
    }
}

/// Parse a `<c:dPt>` into its `(point index, fill)`.
fn parse_data_point(
    reader: &mut Reader<&[u8]>,
    scheme: &SchemeColors<'_>,
) -> Option<(usize, Option<Color>, ChartFillMode)> {
    let mut index: Option<usize> = None;
    let mut fill: Option<Color> = None;
    let mut fill_mode = ChartFillMode::Automatic;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                let properties = parse_shape_properties(reader, b"spPr", scheme);
                fill = fill.or(properties.fill);
                if fill_mode == ChartFillMode::Automatic {
                    fill_mode = properties.fill_mode;
                }
            }
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"idx" =>
            {
                index = xml_util::get_attr_str(e, b"val").and_then(|val| val.parse().ok());
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dPt" => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    Some((index?, fill, fill_mode))
}

/// Read an `<a:ln>` nested anywhere up to `end_tag` into the stroke it states.
///
/// Returns `None` when the element carries no `<a:ln>` at all, which is the
/// difference between "use the automatic line" and "use this one".
fn parse_chart_line(
    reader: &mut Reader<&[u8]>,
    end_tag: &[u8],
    scheme: &SchemeColors<'_>,
) -> ChartLine {
    let mut saw_line: bool = false;
    let mut suppressed: bool = false;
    let mut width_pt: Option<f64> = None;
    let mut color: Option<Color> = None;
    let mut alpha: Option<f64> = None;
    let mut in_line: bool = false;
    let mut in_solid_fill: bool = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"ln" =>
            {
                in_line = true;
                saw_line = true;
                width_pt = width_pt.or_else(|| {
                    xml_util::get_attr_str(e, b"w")
                        .and_then(|w| w.parse::<f64>().ok())
                        .map(|emu| emu / EMU_PER_POINT)
                });
            }
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if in_line && e.local_name().as_ref() == b"noFill" =>
            {
                suppressed = true;
            }
            Ok(Event::Start(ref e)) if in_line && e.local_name().as_ref() == b"solidFill" => {
                in_solid_fill = true;
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                in_solid_fill = false;
            }
            Ok(Event::Start(ref e))
                if in_solid_fill
                    && matches!(e.local_name().as_ref(), b"srgbClr" | b"schemeClr") =>
            {
                let parsed = drawingml::parse_color_from_start(reader, e, scheme);
                if color.is_none() {
                    color = parsed.color;
                    alpha = parsed.alpha;
                }
            }
            Ok(Event::Empty(ref e))
                if in_solid_fill
                    && matches!(e.local_name().as_ref(), b"srgbClr" | b"schemeClr") =>
            {
                let parsed = drawingml::parse_color_from_empty(e, scheme);
                if color.is_none() {
                    color = parsed.color;
                    alpha = parsed.alpha;
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"ln" => {
                in_line = false;
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == end_tag => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    if !saw_line {
        ChartLine::Automatic
    } else if suppressed {
        ChartLine::Suppressed
    } else {
        ChartLine::Explicit {
            width_pt,
            color,
            alpha,
        }
    }
}

/// What a `<c:spPr>` says about the shape it belongs to.
#[derive(Debug, Clone, Copy, Default)]
struct ShapeProperties {
    /// Only shape-level fill declarations affect filled marks. A line's fill
    /// can still supply its colour without changing shape-fill inheritance.
    fill_mode: ChartFillMode,
    /// The first `<a:solidFill>` colour anywhere inside, which for a line
    /// series is the one nested in its `<a:ln>` — that is where a line states
    /// its colour.
    fill: Option<Color>,
    /// The width `<a:ln w="…">` states, in points. `None` when the element
    /// carries no `<a:ln>`, or one that names no usable width.
    line_width_pt: Option<f64>,
    line_geometry: crate::ir::ChartStrokeGeometry,
}

/// Read only the chart style's direct dataPointLine shape properties. Other
/// style roles and nested extension properties cannot supply series defaults.
pub(super) fn parse_series_stroke_defaults(xml: &str) -> crate::ir::ChartStrokeGeometry {
    let mut reader = Reader::from_str(xml);
    let mut path: Vec<Vec<u8>> = Vec::new();
    let mut geometry = crate::ir::ChartStrokeGeometry::default();
    loop {
        let event = reader.read_event();
        let is_start = matches!(&event, Ok(Event::Start(_)));
        match event {
            Ok(Event::Start(ref element)) | Ok(Event::Empty(ref element)) => {
                let name = element.local_name();
                let in_series_shape = path.len() >= 3
                    && path[0] == b"chartStyle"
                    && path[1] == b"dataPointLine"
                    && path[2] == b"spPr";
                if in_series_shape && path.len() == 3 && name.as_ref() == b"ln" {
                    geometry.cap = geometry.cap.or_else(|| read_line_cap(element));
                }
                if in_series_shape
                    && path.len() == 4
                    && path[3] == b"ln"
                    && geometry.join.is_none()
                    && let Some((join, limit)) = read_line_join(element)
                {
                    geometry.join = Some(join);
                    geometry.miter_limit = limit;
                }
                if is_start {
                    path.push(name.as_ref().to_vec());
                }
            }
            Ok(Event::End(_)) => {
                path.pop();
            }
            Ok(Event::Eof) if path.is_empty() => return geometry,
            Ok(Event::Eof) | Err(_) => return crate::ir::ChartStrokeGeometry::default(),
            _ => {}
        }
    }
}

fn read_line_cap(element: &quick_xml::events::BytesStart<'_>) -> Option<crate::ir::LineCap> {
    match xml_util::get_attr_str(element, b"cap").as_deref() {
        Some("flat") => Some(crate::ir::LineCap::Flat),
        Some("rnd") => Some(crate::ir::LineCap::Round),
        Some("sq") => Some(crate::ir::LineCap::Square),
        _ => None,
    }
}

fn read_line_join(
    element: &quick_xml::events::BytesStart<'_>,
) -> Option<(crate::ir::LineJoin, Option<f64>)> {
    match element.local_name().as_ref() {
        b"round" => Some((crate::ir::LineJoin::Round, None)),
        b"bevel" => Some((crate::ir::LineJoin::Bevel, None)),
        b"miter" => {
            // DrawingML percentages use 100000 for 100%. Native Excel's
            // omitted-limit control exports a ratio of 8.
            let limit = xml_util::get_attr_str(element, b"lim")
                .and_then(|value| value.parse::<f64>().ok())
                .filter(|value| value.is_finite() && *value > 0.0)
                .map_or(8.0, |value| value / 100_000.0);
            Some((crate::ir::LineJoin::Miter, Some(limit)))
        }
        _ => None,
    }
}

/// Read a `<c:spPr>` into its fill mode, colour hint, and stroke geometry,
/// consuming up to `end_tag`.
///
/// A chart part declares no theme of its own, so `<a:schemeClr>` resolves
/// against the theme of the document the graphic frame sits in, which the
/// caller supplies (issue #876).
fn parse_shape_properties(
    reader: &mut Reader<&[u8]>,
    end_tag: &[u8],
    scheme: &SchemeColors<'_>,
) -> ShapeProperties {
    let mut in_solid_fill: bool = false;
    let mut in_line: bool = false;
    let mut properties = ShapeProperties::default();

    loop {
        let event = reader.read_event();
        match &event {
            Ok(Event::Start(e)) if e.local_name().as_ref() == b"ln" => in_line = true,
            Ok(Event::End(e)) if e.local_name().as_ref() == b"ln" => in_line = false,
            _ => {}
        }
        match event {
            // The first `<a:ln>` is the shape's own; `w` is in EMU
            // (ECMA-376 §20.1.2.1.15 `ST_LineWidth`), and an `<a:ln>` naming
            // no `w` states nothing about weight (issue #1113).
            //
            // `w="0"` is not a weight either. Office writes it as the
            // "no outline" idiom, always beside `<a:noFill/>` — both series of
            // `office2pdf_repository_workbook.xlsx` and `123233_charts.xlsx`
            // that carry it do — so reading it as a width would stroke nothing
            // where the shape had been drawn at the default all along.
            // Suppression itself is a separate question this does not answer.
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if e.local_name().as_ref() == b"ln" =>
            {
                properties.line_width_pt = properties.line_width_pt.or_else(|| {
                    xml_util::get_attr_str(e, b"w")
                        .and_then(|width| width.parse::<f64>().ok())
                        .map(|emu| emu / EMU_PER_POINT)
                        .filter(|width_pt| *width_pt > 0.0)
                });
                properties.line_geometry.cap =
                    properties.line_geometry.cap.or_else(|| read_line_cap(e));
            }
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if in_line && matches!(e.local_name().as_ref(), b"round" | b"bevel" | b"miter") =>
            {
                if properties.line_geometry.join.is_none()
                    && let Some((join, limit)) = read_line_join(e)
                {
                    properties.line_geometry.join = Some(join);
                    properties.line_geometry.miter_limit = limit;
                }
            }
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                in_solid_fill = true;
                if !in_line && properties.fill_mode == ChartFillMode::Automatic {
                    properties.fill_mode = ChartFillMode::Explicit;
                }
            }
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                if !in_line && e.local_name().as_ref() == b"noFill" =>
            {
                properties.fill_mode = ChartFillMode::Suppressed;
            }
            Ok(Event::Start(ref e))
                if in_solid_fill
                    && matches!(e.local_name().as_ref(), b"srgbClr" | b"schemeClr") =>
            {
                let parsed = drawingml::parse_color_from_start(reader, e, scheme);
                properties.fill = properties.fill.or(parsed.color);
            }
            Ok(Event::Empty(ref e))
                if in_solid_fill
                    && matches!(e.local_name().as_ref(), b"srgbClr" | b"schemeClr") =>
            {
                properties.fill = properties
                    .fill
                    .or(drawingml::parse_color_from_empty(e, scheme).color);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                in_solid_fill = false;
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == end_tag => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    properties
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
                if let Ok(s) = t.xml_content(OOXML_XML_VERSION) {
                    text.push_str(s.as_ref());
                }
            }
            Ok(Event::GeneralRef(ref reference)) if in_v => {
                if let Some(s) = xml_util::decode_general_ref(reference) {
                    text.push_str(&s);
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
    let mut categories: Vec<String> = Vec::new();
    let mut current_text = String::new();
    let mut in_v = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"v" {
                    in_v = true;
                    current_text.clear();
                }
            }
            // One label is one `<c:v>`, not one text event. The reader splits a
            // text node at every entity reference, so a label collected event by
            // event turned `room &amp; board` into the two labels `room ` and
            // ` board` — three of six categories doubled, and every bar past the
            // first labelled with a neighbour's word (issue #1183).
            Ok(Event::Text(ref t)) if in_v => {
                if let Ok(s) = t.xml_content(OOXML_XML_VERSION) {
                    current_text.push_str(s.as_ref());
                }
            }
            Ok(Event::GeneralRef(ref reference)) if in_v => {
                if let Some(s) = xml_util::decode_general_ref(reference) {
                    current_text.push_str(&s);
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"v" => {
                    in_v = false;
                    // A `<c:v>` holding nothing names no label. Excel writes a
                    // blank category as a gap in the `<c:pt idx>` sequence
                    // rather than an empty element, so this only guards a
                    // hand-written part.
                    if !current_text.is_empty() {
                        categories.push(std::mem::take(&mut current_text));
                    }
                }
                b"cat" | b"xVal" => break,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    categories
}

/// Parse numeric values from `<c:val>` or `<c:yVal>`, with the cache's
/// `<c:formatCode>` when it states one.
///
/// `General` is Excel's "no format" and is reported as `None`: applying it
/// would only reformat the number the plain path already prints.
fn parse_value_data(reader: &mut Reader<&[u8]>) -> (Vec<f64>, Option<String>) {
    let mut values = Vec::new();
    let mut format_code: Option<String> = None;
    let mut in_v = false;
    let mut in_format_code = false;
    let mut current_text = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"v" => {
                    in_v = true;
                    current_text.clear();
                }
                b"formatCode" => {
                    in_format_code = true;
                    current_text.clear();
                }
                _ => {}
            },
            Ok(Event::Text(ref t)) if in_v || in_format_code => {
                if let Ok(s) = t.xml_content(OOXML_XML_VERSION) {
                    current_text.push_str(s.as_ref());
                }
            }
            Ok(Event::GeneralRef(ref reference)) if in_v || in_format_code => {
                if let Some(s) = xml_util::decode_general_ref(reference) {
                    current_text.push_str(&s);
                }
            }
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"v" => {
                    in_v = false;
                    if let Ok(v) = current_text.trim().parse::<f64>() {
                        values.push(v);
                    }
                }
                b"formatCode" => {
                    in_format_code = false;
                    let code = current_text.trim();
                    if !code.is_empty() && !code.eq_ignore_ascii_case("General") {
                        format_code = Some(code.to_string());
                    }
                }
                b"val" | b"yVal" => break,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }

    (values, format_code)
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
                            && let Ok(val) = attr.normalized_value(OOXML_XML_VERSION)
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
                            && let Ok(val) = attr.normalized_value(OOXML_XML_VERSION)
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
    crate::parser::xml_util::parse_relationships(rels_xml)
        .into_iter()
        .filter(|entry| {
            entry
                .rel_type
                .as_deref()
                .is_some_and(|rel_type| rel_type.contains("chart"))
        })
        .map(|entry| {
            // Target is relative to word/ directory
            let full_path = if let Some(stripped) = entry.target.strip_prefix('/') {
                stripped.to_string()
            } else {
                format!("word/{}", entry.target)
            };
            (entry.id, full_path)
        })
        .collect()
}

#[cfg(test)]
#[path = "chart_tests.rs"]
mod tests;
