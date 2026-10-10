use std::collections::HashMap;

use crate::ir::Chart;
use crate::parser::chart;

use super::notes::read_zip_text;

pub(in super::super) struct ChartContext {
    charts: HashMap<usize, Vec<Chart>>,
}

impl ChartContext {
    pub(in super::super) fn empty() -> Self {
        Self {
            charts: HashMap::new(),
        }
    }

    pub(in super::super) fn take(&mut self, index: usize) -> Vec<Chart> {
        self.charts.remove(&index).unwrap_or_default()
    }
}

pub(in super::super) fn build_chart_context_from_xml(
    doc_xml: Option<&str>,
    archive: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>,
) -> ChartContext {
    let mut charts: HashMap<usize, Vec<Chart>> = HashMap::new();

    let Some(doc_xml) = doc_xml else {
        return ChartContext { charts };
    };

    let Some(relationships_xml) = read_zip_text(archive, "word/_rels/document.xml.rels") else {
        return ChartContext { charts };
    };

    let chart_references = chart::scan_chart_references(doc_xml);
    let chart_relationships = chart::scan_chart_rels(&relationships_xml);

    // A series that names no fill takes the document theme's accents rather
    // than the renderer's built-in palette (issue #670).
    // The same part settles chart text's face: a chart naming no `a:latin`
    // takes the theme's minor font rather than the engine's default (issue
    // #668).
    let theme_xml: Option<String> = read_zip_text(archive, "word/theme/theme1.xml");
    // The same scheme resolves a series' own `<a:schemeClr>` fill: the chart
    // part declares no theme, so it borrows the document's (issue #876).
    let theme_colors: std::collections::HashMap<String, crate::ir::Color> = theme_xml
        .as_deref()
        .map(crate::parser::drawingml::parse_theme_color_scheme)
        .unwrap_or_default();
    let no_aliases: std::collections::HashMap<String, String> = std::collections::HashMap::new();
    let scheme = crate::parser::drawingml::SchemeColors {
        colors: &theme_colors,
        aliases: &no_aliases,
    };
    let theme_accents: Vec<crate::ir::Color> =
        crate::parser::drawingml::theme_accent_palette(&theme_colors);
    let theme_fonts: crate::parser::drawingml::ThemeFontScheme = theme_xml
        .as_deref()
        .map(crate::parser::drawingml::parse_theme_font_scheme)
        .unwrap_or_default();
    let theme_line_styles: Vec<crate::parser::pptx::ThemeLineStyle> = theme_xml
        .as_deref()
        .map(crate::parser::pptx::parse_theme_line_styles)
        .unwrap_or_default();

    for (body_index, relationship_id) in chart_references {
        if let Some(chart_path) = chart_relationships.get(&relationship_id)
            && let Some(chart_xml) = read_zip_text(archive, chart_path)
            && let Some(mut chart) = crate::parser::chart_style::parse_chart_with_style(
                archive, chart_path, &chart_xml, &scheme,
            )
        {
            chart.theme_accent_colors = theme_accents.clone();
            chart.host = crate::ir::ChartHost::WordProcessing;
            theme_fonts.resolve_chart_text_fonts(&mut chart);
            // The shapes the chart's own drawing part lays over it, which the
            // chart XML can only name through a relationship (issue #1186).
            chart.user_shapes = crate::parser::chart_drawing::load_chart_user_shapes(
                archive,
                chart_path,
                &chart_xml,
                &scheme,
                &theme_fonts,
                &theme_line_styles,
            );
            charts.entry(body_index).or_default().push(chart);
        }
    }

    ChartContext { charts }
}
