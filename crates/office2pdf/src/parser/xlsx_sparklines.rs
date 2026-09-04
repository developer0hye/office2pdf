use std::collections::HashMap;
use std::io::Read;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

#[derive(Debug, Clone, PartialEq)]
pub(crate) struct RawSparklineColor {
    pub(crate) rgb: Option<String>,
    pub(crate) theme: Option<u32>,
    pub(crate) tint: f64,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RawSparkline {
    pub(crate) formula: String,
    pub(crate) destination: String,
}

#[derive(Debug, Clone, PartialEq)]
pub(crate) struct RawSparklineGroup {
    pub(crate) kind: String,
    pub(crate) display_empty_cells_as: String,
    pub(crate) color: RawSparklineColor,
    pub(crate) sparklines: Vec<RawSparkline>,
}

pub(crate) type SheetSparklineGroups = HashMap<String, Vec<RawSparklineGroup>>;
pub(crate) type ResolvedSparklines =
    HashMap<String, HashMap<super::CellPos, crate::ir::SparklineInfo>>;

fn attr_value(reader: &Reader<&[u8]>, element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    element
        .attributes()
        .flatten()
        .find(|attribute| attribute.key.local_name().as_ref() == name)
        .and_then(|attribute| {
            attribute
                .decode_and_unescape_value(reader.decoder())
                .ok()
                .map(|value| value.into_owned())
        })
}

fn raw_color(reader: &Reader<&[u8]>, element: &BytesStart<'_>) -> RawSparklineColor {
    RawSparklineColor {
        rgb: attr_value(reader, element, b"rgb"),
        theme: attr_value(reader, element, b"theme").and_then(|value| value.parse().ok()),
        tint: attr_value(reader, element, b"tint")
            .and_then(|value| value.parse().ok())
            .unwrap_or(0.0),
    }
}

fn parse_worksheet_groups(xml: &str) -> Vec<RawSparklineGroup> {
    let mut groups = Vec::new();
    let mut group: Option<RawSparklineGroup> = None;
    let mut sparkline: Option<RawSparkline> = None;
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"sparklineGroup" => {
                group = Some(RawSparklineGroup {
                    kind: attr_value(&reader, &element, b"type")
                        .unwrap_or_else(|| "line".to_string()),
                    display_empty_cells_as: attr_value(&reader, &element, b"displayEmptyCellsAs")
                        .unwrap_or_else(|| "zero".to_string()),
                    color: RawSparklineColor {
                        rgb: None,
                        theme: None,
                        tint: 0.0,
                    },
                    sparklines: Vec::new(),
                });
            }
            Ok(Event::Start(element) | Event::Empty(element))
                if element.local_name().as_ref() == b"colorSeries" && group.is_some() =>
            {
                group.as_mut().expect("checked above").color = raw_color(&reader, &element);
            }
            Ok(Event::Start(element))
                if element.local_name().as_ref() == b"sparkline" && group.is_some() =>
            {
                sparkline = Some(RawSparkline {
                    formula: String::new(),
                    destination: String::new(),
                });
            }
            Ok(Event::Start(element))
                if matches!(element.local_name().as_ref(), b"f" | b"sqref")
                    && sparkline.is_some() =>
            {
                let local_name = element.local_name().as_ref().to_vec();
                let qualified_name = element.name().to_owned();
                if let Ok(raw) = reader.read_text(qualified_name)
                    && let Ok(text) = quick_xml::escape::unescape(&raw)
                {
                    let sparkline = sparkline.as_mut().expect("checked above");
                    if local_name == b"f" {
                        sparkline.formula = text.into_owned();
                    } else {
                        sparkline.destination = text.into_owned();
                    }
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"sparkline" => {
                if let Some(sparkline) = sparkline.take()
                    && !sparkline.formula.is_empty()
                    && !sparkline.destination.is_empty()
                    && let Some(group) = group.as_mut()
                {
                    group.sparklines.push(sparkline);
                }
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"sparklineGroup" => {
                if let Some(group) = group.take()
                    && !group.sparklines.is_empty()
                {
                    groups.push(group);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    groups
}

pub(crate) fn extract_sparkline_groups(data: &[u8]) -> SheetSparklineGroups {
    let mut archive = match zip::ZipArchive::new(std::io::Cursor::new(data)) {
        Ok(archive) => archive,
        Err(_) => return HashMap::new(),
    };
    let Some(workbook_xml) = super::cond_fmt_raw::read_zip_text(&mut archive, "xl/workbook.xml")
    else {
        return HashMap::new();
    };
    let Some(relationships_xml) =
        super::cond_fmt_raw::read_zip_text(&mut archive, "xl/_rels/workbook.xml.rels")
    else {
        return HashMap::new();
    };
    let relationships = super::cond_fmt_raw::parse_relationships(&relationships_xml);
    let mut result = HashMap::new();
    for (sheet_name, relationship_id) in
        super::cond_fmt_raw::parse_sheet_relationships(&workbook_xml)
    {
        let Some(target) = relationships.get(&relationship_id) else {
            continue;
        };
        let path = super::cond_fmt_raw::worksheet_path(target);
        let Ok(mut file) = archive.by_name(&path) else {
            continue;
        };
        let mut xml = String::new();
        if file.read_to_string(&mut xml).is_err() {
            continue;
        }
        let groups = parse_worksheet_groups(&xml);
        if !groups.is_empty() {
            result.insert(sheet_name, groups);
        }
    }
    result
}

fn formula_range(
    formula: &str,
    destination_sheet: &str,
) -> Option<(String, super::CellPos, super::CellPos)> {
    let formula = formula.trim().trim_start_matches('=');
    let (source_sheet, range) = match formula.rsplit_once('!') {
        Some((sheet, range)) => {
            let sheet = sheet.trim();
            let sheet = sheet
                .strip_prefix('\'')
                .and_then(|sheet| sheet.strip_suffix('\''))
                .unwrap_or(sheet)
                .replace("''", "'");
            (sheet, range)
        }
        None => (destination_sheet.to_string(), formula),
    };
    let (start, end) = range.split_once(':').unwrap_or((range, range));
    Some((
        source_sheet,
        super::parse_cell_ref(start)?,
        super::parse_cell_ref(end)?,
    ))
}

fn resolved_color(
    raw: &RawSparklineColor,
    theme: Option<&umya_spreadsheet::structs::drawing::Theme>,
) -> crate::ir::Color {
    let mut color = umya_spreadsheet::Color::default();
    if let Some(rgb) = raw.rgb.as_deref() {
        color.set_argb(rgb);
    } else {
        // Fall back to the workbook's first accent when the extension omits
        // the normally present `colorSeries` element.
        color.set_theme_index(raw.theme.unwrap_or(4));
    }
    color.set_tint(raw.tint);
    super::xlsx_style::resolve_style_color(&color, theme)
        .unwrap_or_else(|| crate::ir::Color::new(0x5B, 0x9B, 0xD5))
}

fn source_values(
    book: &umya_spreadsheet::Spreadsheet,
    source_sheet: &str,
    start: super::CellPos,
    end: super::CellPos,
    empty_mode: &str,
) -> Option<Vec<Option<f64>>> {
    let sheet = book.get_sheet_by_name(source_sheet)?;
    let (start_col, end_col) = if start.0 <= end.0 {
        (start.0, end.0)
    } else {
        (end.0, start.0)
    };
    let (start_row, end_row) = if start.1 <= end.1 {
        (start.1, end.1)
    } else {
        (end.1, start.1)
    };
    let mut values = Vec::new();
    for row in start_row..=end_row {
        for col in start_col..=end_col {
            values.push(
                sheet
                    .get_cell((col, row))
                    .and_then(|cell| cell.get_value_number()),
            );
        }
    }
    match empty_mode {
        "zero" => {
            for value in &mut values {
                value.get_or_insert(0.0);
            }
        }
        // "span" connects the values on either side of an empty cell.
        "span" => values.retain(Option::is_some),
        // "gap" and unknown future spellings preserve missing points.
        _ => {}
    }
    Some(values)
}

pub(crate) fn resolve_sparklines(
    book: &umya_spreadsheet::Spreadsheet,
    groups_by_sheet: &SheetSparklineGroups,
) -> ResolvedSparklines {
    let mut result = HashMap::new();
    for (destination_sheet, groups) in groups_by_sheet {
        let mut sheet_sparklines = HashMap::new();
        for group in groups.iter().filter(|group| group.kind == "line") {
            let color = resolved_color(&group.color, Some(book.get_theme()));
            for raw in &group.sparklines {
                let Some(destination) = super::parse_cell_ref(&raw.destination) else {
                    continue;
                };
                let Some((source_sheet, start, end)) =
                    formula_range(&raw.formula, destination_sheet)
                else {
                    continue;
                };
                let Some(values) = source_values(
                    book,
                    &source_sheet,
                    start,
                    end,
                    &group.display_empty_cells_as,
                ) else {
                    continue;
                };
                if values.iter().filter(|value| value.is_some()).count() < 2 {
                    continue;
                }
                sheet_sparklines.insert(destination, crate::ir::SparklineInfo { values, color });
            }
        }
        if !sheet_sparklines.is_empty() {
            result.insert(destination_sheet.clone(), sheet_sparklines);
        }
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reported_budget_fixture_exposes_all_ten_x14_sparklines() {
        let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
        let sheets = extract_sparkline_groups(data);
        let groups = sheets
            .get("Monthly college budget")
            .expect("the worksheet x14 extension must be read");
        assert_eq!(groups.len(), 1);

        let group = &groups[0];
        assert_eq!(group.kind, "line");
        assert_eq!(group.display_empty_cells_as, "gap");
        assert_eq!(group.color.theme, Some(6));
        assert!((group.color.tint - -0.499_984_740_745_262).abs() < 1e-12);
        assert_eq!(group.sparklines.len(), 10);
        assert_eq!(
            group.sparklines[0].formula,
            "'Monthly college budget'!C72:N72"
        );
        assert_eq!(group.sparklines[0].destination, "Q72");
        assert!(
            group
                .sparklines
                .iter()
                .any(|sparkline| sparkline.formula.ends_with("C28:N28")
                    && sparkline.destination == "Q28")
        );
    }

    #[test]
    fn reported_budget_sparklines_resolve_cached_values_and_theme_tint() {
        let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
        let book = umya_spreadsheet::reader::xlsx::read_reader(
            std::io::Cursor::new(data.as_slice()),
            true,
        )
        .unwrap();
        let raw = extract_sparkline_groups(data);
        let resolved = resolve_sparklines(&book, &raw);
        let sparkline = &resolved["Monthly college budget"][&(17, 28)];
        assert_eq!(
            sparkline.values,
            vec![
                Some(169.0),
                Some(69.0),
                Some(192.0),
                Some(199.0),
                Some(204.0),
                Some(-771.0),
                Some(124.0),
                Some(154.0),
                Some(-721.0),
                Some(109.0),
                Some(34.0),
                Some(-61.0),
            ]
        );
        assert_eq!(sparkline.color, crate::ir::Color::new(0x29, 0x74, 0x4F));
    }

    #[test]
    fn xlsx_parser_attaches_the_ten_sparklines_to_their_empty_q_cells() {
        use crate::config::ConvertOptions;
        use crate::ir::Page;
        use crate::parser::Parser;

        let data = include_bytes!("../../../../tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx");
        let (document, _) = super::super::XlsxParser
            .parse(data, &ConvertOptions::default())
            .unwrap();
        let page = document
            .pages
            .iter()
            .find_map(|page| match page {
                Page::Sheet(sheet) if sheet.name == "Monthly college budget" => Some(sheet),
                _ => None,
            })
            .expect("the reported worksheet must print");
        let attached = page
            .table
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .filter(|cell| cell.sparkline.is_some())
            .count();
        assert_eq!(attached, 10);

        // The print area begins at A1, and this row has no merges before Q.
        let q28 = &page.table.rows[27].cells[16];
        assert!(
            q28.content.is_empty(),
            "the destination cell carries no text"
        );
        assert_eq!(q28.sparkline.as_ref().unwrap().values[5], Some(-771.0));
    }
}
