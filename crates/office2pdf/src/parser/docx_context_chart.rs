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

    for (body_index, relationship_id) in chart_references {
        if let Some(chart_path) = chart_relationships.get(&relationship_id)
            && let Some(chart_xml) = read_zip_text(archive, chart_path)
            && let Some(chart) = chart::parse_chart_xml(&chart_xml)
        {
            charts.entry(body_index).or_default().push(chart);
        }
    }

    ChartContext { charts }
}
