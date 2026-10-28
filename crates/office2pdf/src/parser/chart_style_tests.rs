use std::io::{Cursor, Write};

use super::*;
use crate::ir::{LineCap, LineJoin};
use crate::parser::drawingml::SchemeColors;

fn styled_chart(
    chart_path: &str,
    relationships: &str,
    style_path: Option<&str>,
    style_line: &str,
    families: &[&str],
) -> Chart {
    let plots = families.iter().map(|family| {
        let (tag, properties) = if *family == "filledRadar" {
            ("radarChart", "<c:radarStyle val=\"filled\"/>")
        } else {
            (*family, "")
        };
        format!("<c:{tag}>{properties}<c:ser><c:val><c:numLit><c:pt idx=\"0\"><c:v>169</c:v></c:pt><c:pt idx=\"1\"><c:v>-771</c:v></c:pt></c:numLit></c:val></c:ser></c:{tag}>")
    }).collect::<String>();
    let xml =
        format!("<c:chartSpace><c:chart><c:plotArea>{plots}</c:plotArea></c:chart></c:chartSpace>");
    let colors = std::collections::HashMap::new();
    let aliases = std::collections::HashMap::new();
    let scheme = SchemeColors {
        colors: &colors,
        aliases: &aliases,
    };
    let (directory, file) = chart_path.rsplit_once('/').unwrap();
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default();
    writer
        .start_file(format!("{directory}/_rels/{file}.rels"), options)
        .unwrap();
    writer.write_all(relationships.as_bytes()).unwrap();
    if let Some(style_path) = style_path {
        writer.start_file(style_path, options).unwrap();
        writer.write_all(format!("<cs:chartStyle><cs:dataPointLine><cs:spPr>{style_line}</cs:spPr></cs:dataPointLine></cs:chartStyle>").as_bytes()).unwrap();
    }
    let mut archive =
        zip::ZipArchive::new(Cursor::new(writer.finish().unwrap().into_inner())).unwrap();
    parse_chart_with_style(&mut archive, chart_path, &xml, &scheme).unwrap()
}

fn relationships(target: &str, kind: &str, mode: &str) -> String {
    format!(
        "<Relationships><Relationship Id=\"style\" Type=\"{kind}\" Target=\"{target}\" TargetMode=\"{mode}\"/></Relationships>"
    )
}

#[test]
fn test_chart_style_resolves_package_relative_and_absolute_parts() {
    for host in ["word", "ppt", "xl"] {
        let path = format!("{host}/charts/chart7.xml");
        let style = format!("{host}/styles/line.xml");
        for target in ["../styles/line.xml".to_owned(), format!("/{style}")] {
            let chart = styled_chart(
                &path,
                &relationships(&target, CHART_STYLE_REL, "Internal"),
                Some(&style),
                "<a:ln cap=\"sq\"><a:bevel/></a:ln>",
                &["lineChart"],
            );
            assert_eq!(
                chart.series[0].line_geometry.cap,
                Some(LineCap::Square),
                "{path}, {target}"
            );
            assert_eq!(chart.series[0].line_geometry.join, Some(LineJoin::Bevel));
        }
    }
}

#[test]
fn test_chart_style_leaves_external_missing_and_unrelated_parts_unresolved() {
    for (kind, mode, part) in [
        (CHART_STYLE_REL, "External", Some("xl/styles/line.xml")),
        (CHART_STYLE_REL, "Internal", None),
        (
            "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle",
            "Internal",
            Some("xl/styles/line.xml"),
        ),
    ] {
        let chart = styled_chart(
            "xl/charts/chart7.xml",
            &relationships("../styles/line.xml", kind, mode),
            part,
            "<a:ln cap=\"rnd\"><a:round/></a:ln>",
            &["lineChart"],
        );
        assert_eq!(chart.series[0].line_geometry.cap, None);
        assert_eq!(chart.series[0].line_geometry.join, None);
        assert_eq!(chart.series[0].values, [169.0, -771.0]);
    }
}

#[test]
fn test_chart_style_applies_only_to_line_families_in_mixed_charts() {
    let families = [
        "barChart",
        "lineChart",
        "scatterChart",
        "radarChart",
        "areaChart",
        "pieChart",
    ];
    let chart = styled_chart(
        "ppt/charts/chart7.xml",
        &relationships("style.xml", CHART_STYLE_REL, "Internal"),
        Some("ppt/charts/style.xml"),
        "<a:ln cap=\"flat\"><a:miter lim=\"600000\"/></a:ln>",
        &families,
    );
    assert_eq!(chart.series.len(), families.len());
    for (index, series) in chart.series.iter().enumerate() {
        let should_inherit = (1..=3).contains(&index);
        assert_eq!(
            series.line_geometry.cap,
            should_inherit.then_some(LineCap::Flat),
            "{}",
            families[index]
        );
        assert_eq!(
            series.line_geometry.join,
            should_inherit.then_some(LineJoin::Miter)
        );
        assert_eq!(
            series.line_geometry.miter_limit,
            should_inherit.then_some(6.0)
        );
    }
}

#[test]
fn test_chart_style_inherited_miter_uses_measured_default_when_limit_is_omitted() {
    let chart = styled_chart(
        "word/charts/chart7.xml",
        &relationships("style.xml", CHART_STYLE_REL, "Internal"),
        Some("word/charts/style.xml"),
        "<a:ln><a:miter/></a:ln>",
        &["lineChart"],
    );
    assert_eq!(chart.series[0].line_geometry.cap, None);
    assert_eq!(chart.series[0].line_geometry.join, Some(LineJoin::Miter));
    assert_eq!(chart.series[0].line_geometry.miter_limit, Some(8.0));
}

#[test]
fn test_chart_style_excludes_3d_line_family() {
    let chart = styled_chart(
        "xl/charts/chart7.xml",
        &relationships("style.xml", CHART_STYLE_REL, "Internal"),
        Some("xl/charts/style.xml"),
        "<a:ln cap=\"rnd\"><a:round/></a:ln>",
        &["line3DChart"],
    );
    assert_eq!(chart.series[0].line_geometry.cap, None);
    assert_eq!(chart.series[0].line_geometry.join, None);
}

#[test]
fn test_chart_style_excludes_filled_radar_family() {
    let chart = styled_chart(
        "xl/charts/chart7.xml",
        &relationships("style.xml", CHART_STYLE_REL, "Internal"),
        Some("xl/charts/style.xml"),
        "<a:ln cap=\"rnd\"><a:round/></a:ln>",
        &["filledRadar"],
    );
    assert_eq!(chart.series[0].line_geometry.cap, None);
    assert_eq!(chart.series[0].line_geometry.join, None);
}
