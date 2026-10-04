#[path = "docx_context_bidi.rs"]
mod bidi;
#[path = "docx_context_chart.rs"]
mod chart;
#[path = "docx_context_columns.rs"]
mod columns;
#[path = "docx_context_contextual_spacing.rs"]
mod contextual_spacing;
#[path = "docx_context_shape.rs"]
mod docx_context_shape;
#[path = "docx_context_drawing.rs"]
mod drawing;
#[path = "docx_context_fields.rs"]
mod fields;
#[path = "docx_context_math.rs"]
mod math;
#[path = "docx_context_notes.rs"]
mod notes;
#[path = "docx_context_page_numbers.rs"]
mod page_numbers;
#[path = "docx_context_paragraph_cursor.rs"]
mod paragraph_cursor;
#[path = "docx_context_paragraph_mark.rs"]
mod paragraph_mark;
#[path = "docx_context_paragraph_shading.rs"]
mod paragraph_shading;
#[path = "docx_context_small_caps.rs"]
mod small_caps;
#[path = "docx_context_table_header.rs"]
mod table_header;
#[path = "docx_context_table_style.rs"]
mod table_style;
#[path = "docx_context_vml.rs"]
mod vml;
#[path = "docx_context_word_wrap.rs"]
mod word_wrap;
#[path = "docx_context_wrap.rs"]
mod wrap;

pub(super) use bidi::BidiContext;
pub(super) use chart::{ChartContext, build_chart_context_from_xml};
pub(super) use columns::{extract_column_layout_from_section_property, scan_column_layouts};
pub(super) use contextual_spacing::{ContextualSpacingContext, ParagraphContextualSpacing};
pub(super) use docx_context_shape::{DrawingShapeContext, WpgDrawingInfo};
pub(super) use drawing::{DrawingTextBoxContext, DrawingTextBoxInfo};
pub(super) use fields::{FieldContext, seq_identifier, toc_caption_identifier, toc_heading_depth};
pub(super) use math::{MathContext, build_math_context_from_xml};
pub(super) use notes::{
    NoteContent, NoteContext, build_note_context_from_xml, is_note_reference_run, read_zip_text,
};
pub(super) use page_numbers::scan_page_numbering;
pub(super) use paragraph_mark::ParagraphMarkContext;
pub(super) use paragraph_shading::{ParagraphShadingContext, scan_style_paragraph_shading};
pub(super) use small_caps::SmallCapsContext;
pub(super) use table_header::TableHeaderContext;
#[cfg(test)]
pub(super) use table_header::scan_table_headers;
pub(super) use table_style::{ResolvedTableStyle, TableStyleContext, apply_table_text_style};
pub(super) use vml::{VmlTextBoxContext, VmlTextBoxInfo};
pub(super) use word_wrap::{WordWrapContext, scan_style_word_wrap};
pub(super) use wrap::{WrapContext, build_wrap_context_from_xml};

/// Bundled conversion contexts threaded through the recursive DOCX call tree.
///
/// Groups the context types that were previously passed as individual
/// parameters, eliminating `#[allow(clippy::too_many_arguments)]` annotations.
pub(super) struct DocxConversionContext {
    pub(super) notes: NoteContext,
    pub(super) wraps: WrapContext,
    pub(super) drawing_text_boxes: DrawingTextBoxContext,
    pub(super) drawing_shapes: DrawingShapeContext,
    pub(super) table_headers: TableHeaderContext,
    pub(super) table_styles: TableStyleContext,
    pub(super) vml_text_boxes: VmlTextBoxContext,
    pub(super) bidi: BidiContext,
    pub(super) small_caps: SmallCapsContext,
    pub(super) paragraph_shading: ParagraphShadingContext,
    pub(super) word_wraps: WordWrapContext,
    /// Each paragraph's `w:contextualSpacing`, which drops `w:spacing` gaps
    /// between paragraphs of the same style (issue #1684).
    pub(super) contextual_spacing: ContextualSpacingContext,
    /// Which paragraph marks a tracked deletion or move removed from the final
    /// document, so those paragraphs merge into the next one (issue #1710).
    pub(super) paragraph_marks: ParagraphMarkContext,
    pub(super) fields: FieldContext,
    /// Whether `word/styles.xml` explicitly defines the default paragraph
    /// style. Decides the East Asian auto space for paragraphs without a
    /// resolvable `w:pStyle` (issue #732).
    pub(super) default_paragraph_style_is_defined: bool,
    /// Whether `word/styles.xml` declares `w:docDefaults/w:pPrDefault`.
    /// Decides the `w:spacing w:after` a paragraph that states none takes:
    /// zero with the element, Word's built-in 8pt without it (issue #1085).
    pub(super) paragraph_property_defaults_are_declared: bool,
}

impl DocxConversionContext {
    /// Build independent cursors for one header/footer part, whose paragraphs
    /// and tables must not consume the document body's pre-scanned entries.
    pub(super) fn for_header_footer_story(
        story_xml: &str,
        styles_xml: Option<&str>,
        default_paragraph_style_is_defined: bool,
        paragraph_property_defaults_are_declared: bool,
    ) -> Self {
        let story_xml: Option<&str> = Some(story_xml);
        Self {
            notes: NoteContext::empty(),
            wraps: build_wrap_context_from_xml(story_xml),
            drawing_text_boxes: DrawingTextBoxContext::from_xml(story_xml),
            drawing_shapes: DrawingShapeContext::from_xml(story_xml),
            table_headers: TableHeaderContext::from_xml(story_xml),
            table_styles: TableStyleContext::from_xml(story_xml, styles_xml),
            vml_text_boxes: VmlTextBoxContext::from_xml(story_xml),
            bidi: BidiContext::from_xml(story_xml),
            small_caps: SmallCapsContext::from_xml(story_xml),
            paragraph_shading: ParagraphShadingContext::from_xml(story_xml),
            word_wraps: WordWrapContext::from_xml(story_xml),
            contextual_spacing: ContextualSpacingContext::from_xml(story_xml, styles_xml),
            paragraph_marks: ParagraphMarkContext::from_xml(story_xml),
            fields: FieldContext::default(),
            default_paragraph_style_is_defined,
            paragraph_property_defaults_are_declared,
        }
    }
}
