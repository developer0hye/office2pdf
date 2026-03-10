#[path = "docx_context_notes.rs"]
mod notes;
#[path = "docx_context_wrap.rs"]
mod wrap;
#[path = "docx_context_drawing.rs"]
mod drawing;
#[path = "docx_context_table_header.rs"]
mod table_header;
#[path = "docx_context_vml.rs"]
mod vml;
#[path = "docx_context_bidi.rs"]
mod bidi;
#[path = "docx_context_small_caps.rs"]
mod small_caps;
#[path = "docx_context_math.rs"]
mod math;
#[path = "docx_context_chart.rs"]
mod chart;
#[path = "docx_context_columns.rs"]
mod columns;

pub(super) use bidi::BidiContext;
pub(super) use chart::{ChartContext, build_chart_context_from_xml};
pub(super) use columns::{extract_column_layout_from_section_property, scan_column_layouts};
pub(super) use drawing::{DrawingTextBoxContext, DrawingTextBoxInfo};
pub(super) use math::{MathContext, build_math_context_from_xml};
pub(super) use notes::{
    NoteContext, build_note_context_from_xml, is_note_reference_run, read_zip_text,
};
pub(super) use small_caps::SmallCapsContext;
pub(super) use table_header::TableHeaderContext;
#[cfg(test)]
pub(super) use table_header::{TableHeaderInfo, scan_table_headers};
pub(super) use vml::{VmlTextBoxContext, VmlTextBoxInfo};
pub(super) use wrap::{WrapContext, build_wrap_context_from_xml};

/// Bundled conversion contexts threaded through the recursive DOCX call tree.
///
/// Groups the 7 context types that were previously passed as individual
/// parameters, eliminating `#[allow(clippy::too_many_arguments)]` annotations.
pub(super) struct DocxConversionContext {
    pub(super) notes: NoteContext,
    pub(super) wraps: WrapContext,
    pub(super) drawing_text_boxes: DrawingTextBoxContext,
    pub(super) table_headers: TableHeaderContext,
    pub(super) vml_text_boxes: VmlTextBoxContext,
    pub(super) bidi: BidiContext,
    pub(super) small_caps: SmallCapsContext,
}
