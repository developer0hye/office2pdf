use crate::ir::ColumnLayout;
use crate::parser::units::twips_to_pt;

pub(in super::super) fn scan_column_layouts(xml: &str) -> Vec<Option<ColumnLayout>> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut layouts: Vec<Option<ColumnLayout>> = Vec::new();

    let mut in_section_properties = false;
    let mut in_columns = false;
    let mut num_columns: u32 = 1;
    let mut spacing_twips: f64 = 720.0;
    let mut equal_width = true;
    let mut column_widths: Vec<f64> = Vec::new();

    let build_layout =
        |num_columns: u32, spacing_twips: f64, equal_width: bool, column_widths: &[f64]| {
            if num_columns < 2 {
                return None;
            }

            Some(ColumnLayout {
                num_columns,
                spacing: twips_to_pt(spacing_twips),
                column_widths: if !equal_width && !column_widths.is_empty() {
                    Some(column_widths.to_vec())
                } else {
                    None
                },
            })
        };

    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(ref element)) => match element.local_name().as_ref()
            {
                b"sectPr" => {
                    in_section_properties = true;
                    num_columns = 1;
                    spacing_twips = 720.0;
                    equal_width = true;
                    column_widths.clear();
                }
                b"cols" if in_section_properties => {
                    in_columns = true;
                    for attribute in element.attributes().flatten() {
                        let key = attribute.key.local_name();
                        if let Ok(value) = attribute.unescape_value() {
                            match key.as_ref() {
                                b"num" => {
                                    if let Ok(parsed) = value.parse::<u32>() {
                                        num_columns = parsed;
                                    }
                                }
                                b"space" => {
                                    if let Ok(parsed) = value.parse::<f64>() {
                                        spacing_twips = parsed;
                                    }
                                }
                                b"equalWidth" => equal_width = value != "0",
                                _ => {}
                            }
                        }
                    }
                }
                b"col" if in_columns => {
                    for attribute in element.attributes().flatten() {
                        if attribute.key.local_name().as_ref() == b"w"
                            && let Ok(value) = attribute.unescape_value()
                            && let Ok(parsed) = value.parse::<f64>()
                        {
                            column_widths.push(twips_to_pt(parsed));
                        }
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Empty(ref element)) => match element.local_name().as_ref()
            {
                b"sectPr" => layouts.push(build_layout(1, 720.0, true, &[])),
                b"cols" if in_section_properties => {
                    in_columns = false;
                    for attribute in element.attributes().flatten() {
                        let key = attribute.key.local_name();
                        if let Ok(value) = attribute.unescape_value() {
                            match key.as_ref() {
                                b"num" => {
                                    if let Ok(parsed) = value.parse::<u32>() {
                                        num_columns = parsed;
                                    }
                                }
                                b"space" => {
                                    if let Ok(parsed) = value.parse::<f64>() {
                                        spacing_twips = parsed;
                                    }
                                }
                                b"equalWidth" => equal_width = value != "0",
                                _ => {}
                            }
                        }
                    }
                }
                b"col" if in_columns => {
                    for attribute in element.attributes().flatten() {
                        if attribute.key.local_name().as_ref() == b"w"
                            && let Ok(value) = attribute.unescape_value()
                            && let Ok(parsed) = value.parse::<f64>()
                        {
                            column_widths.push(twips_to_pt(parsed));
                        }
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
                b"sectPr" => {
                    layouts.push(build_layout(
                        num_columns,
                        spacing_twips,
                        equal_width,
                        &column_widths,
                    ));
                    in_section_properties = false;
                }
                b"cols" => in_columns = false,
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
    }

    layouts
}

pub(in super::super) fn extract_column_layout_from_section_property(
    section_prop: &docx_rs::SectionProperty,
) -> Option<ColumnLayout> {
    if section_prop.columns < 2 {
        return None;
    }

    Some(ColumnLayout {
        num_columns: section_prop.columns as u32,
        spacing: twips_to_pt(section_prop.space as f64),
        column_widths: None,
    })
}
