use super::elements::Block;
use super::style::StyleSheet;

/// Top-level document model produced by parsers and consumed by the renderer.
#[derive(Debug, Clone)]
pub struct Document {
    pub metadata: Metadata,
    pub pages: Vec<Page>,
    pub styles: StyleSheet,
}

/// Document metadata extracted from OOXML `docProps/core.xml` (Dublin Core).
#[derive(Debug, Clone, Default)]
pub struct Metadata {
    pub title: Option<String>,
    pub author: Option<String>,
    pub subject: Option<String>,
    pub description: Option<String>,
    pub created: Option<String>,
    pub modified: Option<String>,
}

/// A page in the document — variant depends on source format.
#[derive(Debug, Clone)]
pub enum Page {
    /// DOCX: flowing text pages.
    Flow(FlowPage),
    /// PPTX: fixed coordinate pages.
    Fixed(FixedPage),
    /// XLSX: table-based pages.
    Table(TablePage),
}

/// Page dimensions.
#[derive(Debug, Clone, Copy)]
pub struct PageSize {
    /// Width in points (1 pt = 1/72 inch).
    pub width: f64,
    /// Height in points.
    pub height: f64,
}

impl Default for PageSize {
    fn default() -> Self {
        Self {
            width: crate::defaults::A4_WIDTH_PT,
            height: crate::defaults::A4_HEIGHT_PT,
        }
    }
}

/// Page margins in points.
#[derive(Debug, Clone, Copy)]
pub struct Margins {
    pub top: f64,
    pub bottom: f64,
    pub left: f64,
    pub right: f64,
}

impl Default for Margins {
    fn default() -> Self {
        Self {
            top: crate::defaults::DEFAULT_MARGIN_PT,
            bottom: crate::defaults::DEFAULT_MARGIN_PT,
            left: crate::defaults::DEFAULT_MARGIN_PT,
            right: crate::defaults::DEFAULT_MARGIN_PT,
        }
    }
}

/// Column layout configuration for multi-column sections.
#[derive(Debug, Clone)]
pub struct ColumnLayout {
    /// Number of columns (must be >= 2 for multi-column layout).
    pub num_columns: u32,
    /// Spacing between columns in points (gutter width).
    pub spacing: f64,
    /// Optional per-column widths in points. When `None`, columns are equal width.
    pub column_widths: Option<Vec<f64>>,
}

/// A flowing-content page (DOCX).
#[derive(Debug, Clone)]
pub struct FlowPage {
    pub size: PageSize,
    pub margins: Margins,
    pub content: Vec<Block>,
    pub header: Option<super::elements::HeaderFooter>,
    pub footer: Option<super::elements::HeaderFooter>,
    /// Optional multi-column layout for the page.
    pub columns: Option<ColumnLayout>,
}

/// A fixed-layout page (PPTX slides).
#[derive(Debug, Clone)]
pub struct FixedPage {
    pub size: PageSize,
    pub elements: Vec<FixedElement>,
    /// Optional background color for the page.
    pub background_color: Option<super::style::Color>,
    /// Optional gradient background (takes precedence over `background_color` when present).
    pub background_gradient: Option<super::elements::GradientFill>,
}

/// An element with fixed position on a page.
#[derive(Debug, Clone)]
pub struct FixedElement {
    /// X position in points from left edge.
    pub x: f64,
    /// Y position in points from top edge.
    pub y: f64,
    /// Width in points.
    pub width: f64,
    /// Height in points.
    pub height: f64,
    /// The content of this element.
    pub kind: FixedElementKind,
}

/// Types of fixed-position elements.
#[derive(Debug, Clone)]
pub enum FixedElementKind {
    TextBox(super::elements::TextBoxData),
    Image(super::elements::ImageData),
    Shape(super::elements::Shape),
    Table(super::elements::Table),
    SmartArt(super::elements::SmartArt),
    Chart(super::elements::Chart),
}

/// A table-based page (XLSX sheets).
#[derive(Debug, Clone)]
pub struct TablePage {
    pub name: String,
    pub size: PageSize,
    pub margins: Margins,
    pub table: super::elements::Table,
    pub header: Option<super::elements::HeaderFooter>,
    pub footer: Option<super::elements::HeaderFooter>,
    /// Charts anchored within this sheet, stored as (anchor_row, chart) where
    /// `anchor_row` is the 1-indexed row number after which the chart is rendered.
    pub charts: Vec<(u32, super::elements::Chart)>,
}

#[cfg(test)]
#[path = "document_tests.rs"]
mod tests;
