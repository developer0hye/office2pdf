use std::cell::Cell;
use std::collections::HashMap;

use super::super::tables::{BorderSideSpec, TableBorderSpec};
use super::super::{Block, Color, TextStyle, parse_hex_color};
use crate::ir::{Alignment, BorderLineStyle, BorderSide, CellBorder, Insets, LineCap, LineJoin};
use crate::parser::units::twips_to_pt;

#[derive(Debug, Clone, Copy, Default, PartialEq)]
struct PartialInsets {
    top: Option<f64>,
    right: Option<f64>,
    bottom: Option<f64>,
    left: Option<f64>,
}

impl PartialInsets {
    fn overlay(self, other: Self) -> Self {
        Self {
            top: other.top.or(self.top),
            right: other.right.or(self.right),
            bottom: other.bottom.or(self.bottom),
            left: other.left.or(self.left),
        }
    }

    fn complete(self) -> Option<Insets> {
        Some(Insets {
            top: self.top?,
            right: self.right?,
            bottom: self.bottom?,
            left: self.left?,
        })
    }

    fn overlay_onto(self, inherited: Insets) -> Insets {
        Insets {
            top: self.top.unwrap_or(inherited.top),
            right: self.right.unwrap_or(inherited.right),
            bottom: self.bottom.unwrap_or(inherited.bottom),
            left: self.left.unwrap_or(inherited.left),
        }
    }
}

const NORMAL_TABLE_STYLE_NAME: &str = "Normal Table";

/// Word's latent built-in `Normal Table` cell margins are 0/108/0/108 twips.
const NORMAL_TABLE_CELL_MARGINS: Insets = Insets {
    top: 0.0,
    right: 5.4,
    bottom: 0.0,
    left: 5.4,
};

#[derive(Debug, Clone, Default)]
struct RegionBorders {
    top: Option<BorderSide>,
    bottom: Option<BorderSide>,
    left: Option<BorderSide>,
    right: Option<BorderSide>,
    inside_h: Option<BorderSide>,
    inside_v: Option<BorderSide>,
}

impl RegionBorders {
    fn overlay(self, other: Self) -> Self {
        Self {
            top: other.top.or(self.top),
            bottom: other.bottom.or(self.bottom),
            left: other.left.or(self.left),
            right: other.right.or(self.right),
            inside_h: other.inside_h.or(self.inside_h),
            inside_v: other.inside_v.or(self.inside_v),
        }
    }
}

#[derive(Debug, Clone, Default)]
struct TableRegionStyle {
    background: Option<Color>,
    text_color: Option<Color>,
    bold: Option<bool>,
    /// `w:caps` from the region's `w:rPr`. A table style is where an invoice
    /// template puts the uppercasing of its header row; the runs themselves
    /// stay mixed-case (issue #845).
    all_caps: Option<bool>,
    /// `w:jc` from the region's `w:pPr`, i.e. how text sits inside the cell.
    /// `w:tblPr/w:jc` is a different property — it places the table box — and
    /// is deliberately not read here.
    alignment: Option<Alignment>,
    /// `w:spacing/@w:after` from the region's `w:pPr`.
    space_after: Option<f64>,
    cell_margins: PartialInsets,
    borders: RegionBorders,
}

impl TableRegionStyle {
    fn overlay(self, other: Self) -> Self {
        Self {
            background: other.background.or(self.background),
            text_color: other.text_color.or(self.text_color),
            bold: other.bold.or(self.bold),
            all_caps: other.all_caps.or(self.all_caps),
            alignment: other.alignment.or(self.alignment),
            space_after: other.space_after.or(self.space_after),
            cell_margins: self.cell_margins.overlay(other.cell_margins),
            borders: self.borders.overlay(other.borders),
        }
    }
}

#[derive(Debug, Clone, Default)]
struct TableStyleDefinition {
    name: Option<String>,
    base: TableRegionStyle,
    first_row: TableRegionStyle,
    last_row: TableRegionStyle,
    first_column: TableRegionStyle,
    last_column: TableRegionStyle,
    band1_horizontal: TableRegionStyle,
    band2_horizontal: TableRegionStyle,
    band1_vertical: TableRegionStyle,
    band2_vertical: TableRegionStyle,
}

#[derive(Debug, Clone, Copy)]
struct TableLook {
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
    horizontal_banding: bool,
    vertical_banding: bool,
}

impl Default for TableLook {
    fn default() -> Self {
        Self {
            first_row: false,
            last_row: false,
            first_column: false,
            last_column: false,
            horizontal_banding: true,
            vertical_banding: true,
        }
    }
}

#[derive(Debug, Clone, Default)]
struct TableStyleApplication {
    style_id: Option<String>,
    look: TableLook,
}

struct TableApplicationScanState {
    application_index: usize,
    in_properties: bool,
}

pub(in super::super) struct TableStyleContext {
    styles: HashMap<String, TableStyleDefinition>,
    default_style_id: Option<String>,
    applications: Vec<TableStyleApplication>,
    cursor: Cell<usize>,
}

#[derive(Debug, Clone, Default)]
pub(in super::super) struct ResolvedTableCellStyle {
    pub(in super::super) background: Option<Color>,
    pub(in super::super) text_color: Option<Color>,
    pub(in super::super) bold: Option<bool>,
    pub(in super::super) all_caps: Option<bool>,
    pub(in super::super) alignment: Option<Alignment>,
    pub(in super::super) padding: Option<Insets>,
    pub(in super::super) border: Option<CellBorder>,
}

pub(in super::super) struct ResolvedTableStyle {
    definition: TableStyleDefinition,
    look: TableLook,
    default_cell_padding: Option<Insets>,
}

impl TableStyleContext {
    pub(in super::super) fn from_xml(document_xml: Option<&str>, styles_xml: Option<&str>) -> Self {
        let scanned = styles_xml.map(scan_table_styles).unwrap_or_default();
        Self {
            styles: scanned.styles,
            default_style_id: scanned.default_style_id,
            applications: document_xml
                .map(scan_table_style_applications)
                .unwrap_or_default(),
            cursor: Cell::new(0),
        }
    }

    pub(in super::super) fn consume_next(&self) -> Option<ResolvedTableStyle> {
        let index = self.cursor.get();
        self.cursor.set(index + 1);
        let application = self.applications.get(index)?;
        let style_id = application
            .style_id
            .as_deref()
            .or(self.default_style_id.as_deref())?;
        let definition = self.styles.get(style_id)?.clone();
        let default_style_margins = self
            .default_style_id
            .as_deref()
            .and_then(|default_id| self.styles.get(default_id))
            .map(|default| resolve_table_style_cell_margins(default, Insets::default()))
            .unwrap_or_default();
        let default_cell_padding = if self.default_style_id.as_deref() == Some(style_id) {
            default_style_margins
        } else {
            resolve_table_style_cell_margins(&definition, default_style_margins)
        };
        Some(ResolvedTableStyle {
            definition,
            look: application.look,
            default_cell_padding: Some(default_cell_padding),
        })
    }
}

fn resolve_table_style_cell_margins(
    definition: &TableStyleDefinition,
    inherited: Insets,
) -> Insets {
    if definition
        .name
        .as_deref()
        .is_some_and(|name| name.eq_ignore_ascii_case(NORMAL_TABLE_STYLE_NAME))
    {
        return definition
            .base
            .cell_margins
            .complete()
            .unwrap_or(NORMAL_TABLE_CELL_MARGINS);
    }

    definition.base.cell_margins.overlay_onto(inherited)
}

impl ResolvedTableStyle {
    pub(in super::super) fn default_cell_padding(&self) -> Option<Insets> {
        self.default_cell_padding
    }

    pub(in super::super) fn paragraph_space_after(
        &self,
        row_index: usize,
        row_count: usize,
        column_index: usize,
        column_span: usize,
        column_count: usize,
    ) -> Option<f64> {
        self.resolved_cell_region(
            row_index,
            row_count,
            column_index,
            column_span,
            column_count,
        )
        .space_after
    }

    pub(in super::super) fn cell_style(
        &self,
        row_index: usize,
        row_count: usize,
        column_index: usize,
        column_span: usize,
        column_count: usize,
        direct_borders: &TableBorderSpec,
    ) -> ResolvedTableCellStyle {
        let region: TableRegionStyle = self.resolved_cell_region(
            row_index,
            row_count,
            column_index,
            column_span,
            column_count,
        );
        // Resolve the cell's border sides. The base region draws the table
        // grid (outer edges on boundary cells, insideH/insideV on interior
        // edges); an active special region's explicit sides then override
        // the matching edges of its own cells (e.g. the firstRow bottom
        // border lands on the header row's bottom edge).
        //
        // The table's own `w:tblBorders` sits between the two: it is direct
        // formatting, so it outranks the style's grid on every side it states
        // — including a `w:val="none"`, which means "draw nothing" rather than
        // "say nothing" and so must survive as a suppression (issue #931).
        let base = &self.definition.base.borders;
        let grid = |at_boundary: bool, outer: &Option<BorderSide>, inside: &Option<BorderSide>| {
            if at_boundary {
                outer.clone()
            } else {
                inside.clone()
            }
        };
        let at_top = row_index == 0;
        let at_bottom = row_index + 1 == row_count;
        let at_left = column_index == 0;
        let at_right = column_index + column_span >= column_count;
        let direct = |stated: &BorderSideSpec, from_style: Option<BorderSide>| {
            if stated.is_stated() {
                stated.drawn()
            } else {
                from_style
            }
        };
        let mut top = direct(
            direct_borders.horizontal(at_top, &direct_borders.top),
            grid(at_top, &base.top, &base.inside_h),
        );
        let mut bottom = direct(
            direct_borders.horizontal(at_bottom, &direct_borders.bottom),
            grid(at_bottom, &base.bottom, &base.inside_h),
        );
        let mut left = direct(
            direct_borders.vertical(at_left, &direct_borders.left),
            grid(at_left, &base.left, &base.inside_v),
        );
        let mut right = direct(
            direct_borders.vertical(at_right, &direct_borders.right),
            grid(at_right, &base.right, &base.inside_v),
        );

        let mut override_edges = |borders: &RegionBorders| {
            if let Some(side) = &borders.top {
                top = Some(side.clone());
            }
            if let Some(side) = &borders.bottom {
                bottom = Some(side.clone());
            }
            if let Some(side) = &borders.left {
                left = Some(side.clone());
            }
            if let Some(side) = &borders.right {
                right = Some(side.clone());
            }
        };
        if self.look.first_column && column_index == 0 {
            override_edges(&self.definition.first_column.borders);
        }
        if self.look.last_column && column_index + column_span == column_count {
            override_edges(&self.definition.last_column.borders);
        }
        if self.look.first_row && row_index == 0 {
            override_edges(&self.definition.first_row.borders);
        }
        if self.look.last_row && row_index + 1 == row_count {
            override_edges(&self.definition.last_row.borders);
        }
        let border = (top.is_some() || bottom.is_some() || left.is_some() || right.is_some())
            .then_some(CellBorder {
                top,
                bottom,
                left,
                right,
            });

        ResolvedTableCellStyle {
            background: region.background,
            text_color: region.text_color,
            bold: region.bold,
            all_caps: region.all_caps,
            alignment: region.alignment,
            padding: (region.cell_margins != self.definition.base.cell_margins)
                .then(|| region.cell_margins.complete())
                .flatten(),
            border,
        }
    }

    fn resolved_cell_region(
        &self,
        row_index: usize,
        row_count: usize,
        column_index: usize,
        column_span: usize,
        column_count: usize,
    ) -> TableRegionStyle {
        let mut region = self.definition.base.clone();
        if self.look.horizontal_banding {
            let band_index = row_index.saturating_sub(usize::from(self.look.first_row));
            region = region.overlay(if band_index.is_multiple_of(2) {
                self.definition.band1_horizontal.clone()
            } else {
                self.definition.band2_horizontal.clone()
            });
        }
        if self.look.vertical_banding {
            let band_index = column_index.saturating_sub(usize::from(self.look.first_column));
            region = region.overlay(if band_index.is_multiple_of(2) {
                self.definition.band1_vertical.clone()
            } else {
                self.definition.band2_vertical.clone()
            });
        }
        if self.look.first_row && row_index == 0 {
            region = region.overlay(self.definition.first_row.clone());
        }
        if self.look.last_row && row_index + 1 == row_count {
            region = region.overlay(self.definition.last_row.clone());
        }
        if self.look.first_column && column_index == 0 {
            region = region.overlay(self.definition.first_column.clone());
        }
        if self.look.last_column && column_index + column_span == column_count {
            region = region.overlay(self.definition.last_column.clone());
        }
        region
    }
}

pub(in super::super) fn apply_table_text_style(
    blocks: &mut [Block],
    region: &ResolvedTableCellStyle,
) {
    for block in blocks {
        apply_text_style(block, region);
    }
}

fn apply_text_style(block: &mut Block, region: &ResolvedTableCellStyle) {
    if let Block::Paragraph(paragraph) = block {
        // Direct formatting wins, so the style only fills what the paragraph
        // left unsaid (issue #845).
        if paragraph.style.alignment.is_none() {
            paragraph.style.alignment = region.alignment;
        }
        for run in &mut paragraph.runs {
            let mut style: TextStyle = run.style.clone();
            if style.color.is_none() {
                style.color = region.text_color;
            }
            if style.bold.is_none() {
                style.bold = region.bold;
            }
            if style.all_caps.is_none() {
                style.all_caps = region.all_caps;
            }
            run.style = style;
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum TableStyleRegion {
    Base,
    FirstRow,
    LastRow,
    FirstColumn,
    LastColumn,
    Band1Horizontal,
    Band2Horizontal,
    Band1Vertical,
    Band2Vertical,
    Unsupported,
}

fn region_mut(
    definition: &mut TableStyleDefinition,
    region: TableStyleRegion,
) -> Option<&mut TableRegionStyle> {
    match region {
        TableStyleRegion::Base => Some(&mut definition.base),
        TableStyleRegion::FirstRow => Some(&mut definition.first_row),
        TableStyleRegion::LastRow => Some(&mut definition.last_row),
        TableStyleRegion::FirstColumn => Some(&mut definition.first_column),
        TableStyleRegion::LastColumn => Some(&mut definition.last_column),
        TableStyleRegion::Band1Horizontal => Some(&mut definition.band1_horizontal),
        TableStyleRegion::Band2Horizontal => Some(&mut definition.band2_horizontal),
        TableStyleRegion::Band1Vertical => Some(&mut definition.band1_vertical),
        TableStyleRegion::Band2Vertical => Some(&mut definition.band2_vertical),
        TableStyleRegion::Unsupported => None,
    }
}

/// Parse one border side element (`<w:top w:val w:sz w:color/>`).
/// Widths are eighths of a point; nil/none sides are skipped.
fn parse_border_side(element: &quick_xml::events::BytesStart<'_>) -> Option<BorderSide> {
    let val = attribute_value(element, b"val").unwrap_or_default();
    if val == "nil" || val == "none" || val.is_empty() {
        return None;
    }
    let width: f64 = attribute_value(element, b"sz")
        .and_then(|value| value.parse::<f64>().ok())
        .map(|eighths| eighths / 8.0)
        .unwrap_or(0.5);
    let color: Color = attribute_value(element, b"color")
        .filter(|value| value != "auto")
        .and_then(|value| parse_hex_color(&value))
        .unwrap_or(Color::new(0, 0, 0));
    let style = match val.as_str() {
        "dashed" | "dashSmallGap" => BorderLineStyle::Dashed,
        "dotted" => BorderLineStyle::Dotted,
        "dotDash" => BorderLineStyle::DashDot,
        "dotDotDash" => BorderLineStyle::DashDotDot,
        "double" | "doubleWave" => BorderLineStyle::Double,
        _ => BorderLineStyle::Solid,
    };
    Some(BorderSide {
        width,
        color,
        style,
        join: LineJoin::Round,
        cap: LineCap::Flat,
    })
}

#[derive(Default)]
struct ScannedTableStyles {
    styles: HashMap<String, TableStyleDefinition>,
    default_style_id: Option<String>,
}

fn scan_table_styles(xml: &str) -> ScannedTableStyles {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut styles: HashMap<String, TableStyleDefinition> = HashMap::new();
    let mut default_style_id: Option<String> = None;
    let mut current_style_id: Option<String> = None;
    let mut current_definition = TableStyleDefinition::default();
    let mut current_region = TableStyleRegion::Base;
    let mut in_cell_properties = false;
    let mut in_run_properties = false;
    let mut in_paragraph_properties = false;
    let mut in_borders = false;
    let mut in_cell_margins = false;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(quick_xml::events::Event::Start(ref element)) => {
                match element.local_name().as_ref() {
                    b"style" if attribute_value(element, b"type").as_deref() == Some("table") => {
                        current_style_id = attribute_value(element, b"styleId");
                        if boolean_attribute(element, b"default").unwrap_or(false) {
                            default_style_id.clone_from(&current_style_id);
                        }
                        current_definition = TableStyleDefinition::default();
                        current_region = TableStyleRegion::Base;
                    }
                    b"tblStylePr" if current_style_id.is_some() => {
                        current_region = attribute_value(element, b"type")
                            .as_deref()
                            .map(parse_region)
                            .unwrap_or(TableStyleRegion::Unsupported);
                    }
                    b"tcPr" if current_style_id.is_some() => in_cell_properties = true,
                    b"rPr" if current_style_id.is_some() => in_run_properties = true,
                    b"pPr" if current_style_id.is_some() => in_paragraph_properties = true,
                    b"tblBorders" | b"tcBorders" if current_style_id.is_some() => in_borders = true,
                    b"tblCellMar" | b"tcMar" if current_style_id.is_some() => {
                        in_cell_margins = true
                    }
                    _ => {}
                }
                if current_style_id.is_some() && element.local_name().as_ref() == b"name" {
                    current_definition.name = attribute_value(element, b"val");
                }
                apply_style_element(
                    element,
                    &mut current_definition,
                    current_region,
                    StyleElementScope {
                        in_cell_properties,
                        in_run_properties,
                        in_paragraph_properties,
                        in_borders,
                        in_cell_margins,
                    },
                );
            }
            Ok(quick_xml::events::Event::Empty(ref element)) => {
                if current_style_id.is_some() && element.local_name().as_ref() == b"name" {
                    current_definition.name = attribute_value(element, b"val");
                }
                apply_style_element(
                    element,
                    &mut current_definition,
                    current_region,
                    StyleElementScope {
                        in_cell_properties,
                        in_run_properties,
                        in_paragraph_properties,
                        in_borders,
                        in_cell_margins,
                    },
                );
            }
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
                b"tcPr" => in_cell_properties = false,
                b"rPr" => in_run_properties = false,
                b"pPr" => in_paragraph_properties = false,
                b"tblBorders" | b"tcBorders" => in_borders = false,
                b"tblCellMar" | b"tcMar" => in_cell_margins = false,
                b"tblStylePr" => current_region = TableStyleRegion::Base,
                b"style" => {
                    if let Some(style_id) = current_style_id.take() {
                        styles.insert(style_id, current_definition.clone());
                    }
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buffer.clear();
    }

    ScannedTableStyles {
        styles,
        default_style_id,
    }
}

/// Which property container the scanner is currently inside. `w:jc` appears
/// in both `w:pPr` and `w:tblPr` with unrelated meanings, so the containers
/// have to be distinguished rather than merged.
#[derive(Debug, Clone, Copy)]
struct StyleElementScope {
    in_cell_properties: bool,
    in_run_properties: bool,
    in_paragraph_properties: bool,
    in_borders: bool,
    in_cell_margins: bool,
}

fn apply_style_element(
    element: &quick_xml::events::BytesStart<'_>,
    definition: &mut TableStyleDefinition,
    region: TableStyleRegion,
    scope: StyleElementScope,
) {
    let StyleElementScope {
        in_cell_properties,
        in_run_properties,
        in_paragraph_properties,
        in_borders,
        in_cell_margins,
    } = scope;
    let Some(target) = region_mut(definition, region) else {
        return;
    };
    if in_borders {
        let side_slot: Option<&mut Option<BorderSide>> = match element.local_name().as_ref() {
            b"top" => Some(&mut target.borders.top),
            b"bottom" => Some(&mut target.borders.bottom),
            b"left" | b"start" => Some(&mut target.borders.left),
            b"right" | b"end" => Some(&mut target.borders.right),
            b"insideH" => Some(&mut target.borders.inside_h),
            b"insideV" => Some(&mut target.borders.inside_v),
            _ => None,
        };
        if let Some(slot) = side_slot {
            *slot = parse_border_side(element);
            return;
        }
    }
    if in_cell_margins {
        let value = attribute_value(element, b"w")
            .and_then(|value| value.parse::<f64>().ok())
            .map(crate::parser::units::twips_to_pt);
        let slot = match element.local_name().as_ref() {
            b"top" => Some(&mut target.cell_margins.top),
            b"right" | b"end" => Some(&mut target.cell_margins.right),
            b"bottom" => Some(&mut target.cell_margins.bottom),
            b"left" | b"start" => Some(&mut target.cell_margins.left),
            _ => None,
        };
        if let Some(slot) = slot {
            *slot = value;
            return;
        }
    }
    match element.local_name().as_ref() {
        b"shd" if in_cell_properties => {
            target.background = attribute_value(element, b"fill")
                .filter(|value| value != "auto")
                .and_then(|value| parse_hex_color(&value));
        }
        b"color" if in_run_properties => {
            target.text_color = attribute_value(element, b"val")
                .filter(|value| value != "auto")
                .and_then(|value| parse_hex_color(&value));
        }
        b"b" | b"bCs" if in_run_properties => {
            target.bold = Some(on_off_element_is_enabled(element));
        }
        b"caps" if in_run_properties => {
            target.all_caps = Some(on_off_element_is_enabled(element));
        }
        b"jc" if in_paragraph_properties => {
            target.alignment = attribute_value(element, b"val")
                .as_deref()
                .and_then(super::super::text::parse_alignment);
        }
        b"spacing" if in_paragraph_properties => {
            target.space_after = attribute_value(element, b"after")
                .and_then(|value| value.parse::<f64>().ok())
                .map(twips_to_pt)
                .or(target.space_after);
        }
        _ => {}
    }
}

fn parse_region(value: &str) -> TableStyleRegion {
    match value {
        "firstRow" => TableStyleRegion::FirstRow,
        "lastRow" => TableStyleRegion::LastRow,
        "firstCol" => TableStyleRegion::FirstColumn,
        "lastCol" => TableStyleRegion::LastColumn,
        "band1Horz" => TableStyleRegion::Band1Horizontal,
        "band2Horz" => TableStyleRegion::Band2Horizontal,
        "band1Vert" => TableStyleRegion::Band1Vertical,
        "band2Vert" => TableStyleRegion::Band2Vertical,
        _ => TableStyleRegion::Unsupported,
    }
}

fn scan_table_style_applications(xml: &str) -> Vec<TableStyleApplication> {
    let mut reader = quick_xml::Reader::from_str(xml);
    let mut buffer: Vec<u8> = Vec::new();
    let mut applications: Vec<TableStyleApplication> = Vec::new();
    let mut stack: Vec<TableApplicationScanState> = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(quick_xml::events::Event::Start(ref element)) => {
                match element.local_name().as_ref() {
                    b"tbl" => {
                        applications.push(TableStyleApplication::default());
                        stack.push(TableApplicationScanState {
                            application_index: applications.len() - 1,
                            in_properties: false,
                        });
                    }
                    b"tblPr" => {
                        if let Some(state) = stack.last_mut() {
                            state.in_properties = true;
                        }
                    }
                    _ => {}
                }
                apply_table_application_element(element, &mut applications, stack.last());
            }
            Ok(quick_xml::events::Event::Empty(ref element)) => {
                apply_table_application_element(element, &mut applications, stack.last());
            }
            Ok(quick_xml::events::Event::End(ref element)) => match element.local_name().as_ref() {
                b"tblPr" => {
                    if let Some(state) = stack.last_mut() {
                        state.in_properties = false;
                    }
                }
                b"tbl" => {
                    stack.pop();
                }
                _ => {}
            },
            Ok(quick_xml::events::Event::Eof) | Err(_) => break,
            _ => {}
        }
        buffer.clear();
    }

    applications
}

fn apply_table_application_element(
    element: &quick_xml::events::BytesStart<'_>,
    applications: &mut [TableStyleApplication],
    state: Option<&TableApplicationScanState>,
) {
    let Some(state) = state else {
        return;
    };
    if !state.in_properties {
        return;
    }
    let application = &mut applications[state.application_index];
    match element.local_name().as_ref() {
        b"tblStyle" => application.style_id = attribute_value(element, b"val"),
        b"tblLook" => {
            application.look.first_row = boolean_attribute(element, b"firstRow").unwrap_or(false);
            application.look.last_row = boolean_attribute(element, b"lastRow").unwrap_or(false);
            application.look.first_column =
                boolean_attribute(element, b"firstColumn").unwrap_or(false);
            application.look.last_column =
                boolean_attribute(element, b"lastColumn").unwrap_or(false);
            application.look.horizontal_banding =
                !boolean_attribute(element, b"noHBand").unwrap_or(false);
            application.look.vertical_banding =
                !boolean_attribute(element, b"noVBand").unwrap_or(false);
        }
        _ => {}
    }
}

fn attribute_value(element: &quick_xml::events::BytesStart<'_>, name: &[u8]) -> Option<String> {
    element.attributes().flatten().find_map(|attribute| {
        (attribute.key.local_name().as_ref() == name)
            .then(|| String::from_utf8_lossy(attribute.value.as_ref()).into_owned())
    })
}

fn boolean_attribute(element: &quick_xml::events::BytesStart<'_>, name: &[u8]) -> Option<bool> {
    attribute_value(element, name).map(|value| {
        !value.eq_ignore_ascii_case("0")
            && !value.eq_ignore_ascii_case("false")
            && !value.eq_ignore_ascii_case("off")
    })
}

fn on_off_element_is_enabled(element: &quick_xml::events::BytesStart<'_>) -> bool {
    boolean_attribute(element, b"val").unwrap_or(true)
}

#[cfg(test)]
mod tests {
    use super::*;

    const STYLES_XML: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="table" w:styleId="DarkGrid">
        <w:tblPr>
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="FFFFFF"/>
            <w:left w:val="single" w:sz="4" w:color="FFFFFF"/>
            <w:bottom w:val="single" w:sz="4" w:color="FFFFFF"/>
            <w:right w:val="single" w:sz="4" w:color="FFFFFF"/>
            <w:insideH w:val="single" w:sz="4" w:color="FFFFFF"/>
            <w:insideV w:val="single" w:sz="4" w:color="FFFFFF"/>
          </w:tblBorders>
        </w:tblPr>
        <w:tcPr><w:shd w:val="clear" w:fill="404040"/></w:tcPr>
        <w:tblStylePr w:type="firstRow">
          <w:tcPr>
            <w:tcBorders>
              <w:bottom w:val="double" w:sz="8" w:color="FF0000"/>
            </w:tcBorders>
            <w:shd w:val="clear" w:fill="000000"/>
          </w:tcPr>
        </w:tblStylePr>
      </w:style>
    </w:styles>"#;

    const DOCUMENT_XML: &str = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:body><w:tbl><w:tblPr>
        <w:tblStyle w:val="DarkGrid"/>
        <w:tblLook w:firstRow="1" w:noHBand="1" w:noVBand="1"/>
      </w:tblPr></w:tbl></w:body>
    </w:document>"#;

    #[test]
    fn test_table_style_borders_resolve_per_cell() {
        let context = TableStyleContext::from_xml(Some(DOCUMENT_XML), Some(STYLES_XML));
        let resolved = context.consume_next().expect("style application");

        // Interior cell: white grid on all sides from tblBorders.
        let interior = resolved.cell_style(1, 3, 1, 1, 3, &TableBorderSpec::default());
        let border = interior.border.expect("interior cell gets style borders");
        for side in [&border.top, &border.bottom, &border.left, &border.right] {
            let side = side.as_ref().expect("all sides bordered");
            assert_eq!(side.color, Color::new(0xFF, 0xFF, 0xFF));
            assert_eq!(side.width, 0.5, "w:sz=4 eighths = 0.5pt");
        }

        // First-row cell: red double bottom border from the region override.
        let header = resolved.cell_style(0, 3, 1, 1, 3, &TableBorderSpec::default());
        let border = header.border.expect("header cell gets style borders");
        let bottom = border.bottom.as_ref().expect("bottom side");
        assert_eq!(bottom.color, Color::new(0xFF, 0, 0));
        assert_eq!(bottom.width, 1.0, "w:sz=8 eighths = 1pt");
        assert_eq!(
            bottom.style,
            crate::ir::BorderLineStyle::Double,
            "double border style survives"
        );
        assert_eq!(header.background, Some(Color::new(0, 0, 0)));

        // Boundary edges use the outer borders (still white here).
        let top_left = resolved.cell_style(1, 3, 0, 1, 3, &TableBorderSpec::default());
        let border = top_left.border.expect("boundary cell borders");
        assert!(border.left.is_some());
    }

    /// A table whose own `w:tblBorders` states `w:val="none"` on every side
    /// asks for no borders, and direct formatting outranks the style's grid.
    /// Reading `none` as "nothing stated" left the invoice of #841 drawing six
    /// grey rules the document turns off (issue #931).
    #[test]
    fn a_stated_none_suppresses_the_styles_grid() {
        let context = TableStyleContext::from_xml(Some(DOCUMENT_XML), Some(STYLES_XML));
        let resolved = context.consume_next().expect("style application");
        let all_none: TableBorderSpec = TableBorderSpec {
            top: BorderSideSpec::Suppressed,
            bottom: BorderSideSpec::Suppressed,
            left: BorderSideSpec::Suppressed,
            right: BorderSideSpec::Suppressed,
            inside_h: BorderSideSpec::Suppressed,
            inside_v: BorderSideSpec::Suppressed,
        };

        // An interior cell keeps nothing at all.
        let interior = resolved.cell_style(1, 3, 1, 1, 3, &all_none);
        assert!(
            interior.border.is_none(),
            "a suppressed grid draws no side: {:?}",
            interior.border
        );
        // A boundary cell likewise — the outer sides are suppressed too.
        let corner = resolved.cell_style(2, 3, 0, 1, 3, &all_none);
        assert!(corner.border.is_none(), "{:?}", corner.border);

        // The conditional region still rules its own edge: `firstRow` states a
        // red double bottom border, which direct table-level formatting does
        // not reach.
        let header = resolved.cell_style(0, 3, 1, 1, 3, &all_none);
        let border = header.border.expect("the firstRow region survives");
        let bottom = border.bottom.as_ref().expect("its bottom side");
        assert_eq!(bottom.color, Color::new(0xFF, 0, 0));
        assert!(
            border.top.is_none() && border.left.is_none() && border.right.is_none(),
            "only the side the region states survives: {border:?}"
        );
    }

    /// An unstated side still inherits, and a side the table states outright
    /// replaces the style's rather than being ignored (issue #931).
    #[test]
    fn an_unstated_side_inherits_while_a_stated_one_replaces() {
        let context = TableStyleContext::from_xml(Some(DOCUMENT_XML), Some(STYLES_XML));
        let resolved = context.consume_next().expect("style application");
        let only_inside_h: TableBorderSpec = TableBorderSpec {
            inside_h: BorderSideSpec::Stated(BorderSide {
                width: 2.0,
                color: Color::new(0, 0, 0xFF),
                style: crate::ir::BorderLineStyle::Solid,
                join: LineJoin::Round,
                cap: LineCap::Flat,
            }),
            ..TableBorderSpec::default()
        };
        let interior = resolved.cell_style(1, 3, 1, 1, 3, &only_inside_h);
        let border = interior.border.expect("borders resolved");
        let top = border.top.as_ref().expect("insideH reaches the top edge");
        assert_eq!(
            top.color,
            Color::new(0, 0, 0xFF),
            "the table's own side wins"
        );
        assert_eq!(top.width, 2.0);
        let left = border.left.as_ref().expect("the vertical side is unstated");
        assert_eq!(
            left.color,
            Color::new(0xFF, 0xFF, 0xFF),
            "an unstated side still inherits the style's"
        );
    }
}
