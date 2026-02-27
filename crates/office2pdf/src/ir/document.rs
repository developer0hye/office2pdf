use super::elements::Block;
use super::style::StyleSheet;

/// Top-level document model produced by parsers and consumed by the renderer.
#[derive(Debug, Clone)]
pub struct Document {
    pub metadata: Metadata,
    pub pages: Vec<Page>,
    pub styles: StyleSheet,
}

/// Document metadata.
#[derive(Debug, Clone, Default)]
pub struct Metadata {
    pub title: Option<String>,
    pub author: Option<String>,
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
        // A4 in points
        Self {
            width: 595.28,
            height: 841.89,
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
        // 1 inch = 72 points
        Self {
            top: 72.0,
            bottom: 72.0,
            left: 72.0,
            right: 72.0,
        }
    }
}

/// A flowing-content page (DOCX).
#[derive(Debug, Clone)]
pub struct FlowPage {
    pub size: PageSize,
    pub margins: Margins,
    pub content: Vec<Block>,
    pub header: Option<super::elements::HeaderFooter>,
    pub footer: Option<super::elements::HeaderFooter>,
}

/// A fixed-layout page (PPTX slides).
#[derive(Debug, Clone)]
pub struct FixedPage {
    pub size: PageSize,
    pub elements: Vec<FixedElement>,
    /// Optional background color for the page.
    pub background_color: Option<super::style::Color>,
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
    TextBox(Vec<Block>),
    Image(super::elements::ImageData),
    Shape(super::elements::Shape),
    Table(super::elements::Table),
    SmartArt(super::elements::SmartArt),
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
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_page_size_is_a4() {
        let size = PageSize::default();
        assert!((size.width - 595.28).abs() < 0.01);
        assert!((size.height - 841.89).abs() < 0.01);
    }

    #[test]
    fn test_default_margins_are_one_inch() {
        let margins = Margins::default();
        assert!((margins.top - 72.0).abs() < 0.01);
        assert!((margins.left - 72.0).abs() < 0.01);
    }

    #[test]
    fn test_fixed_page_background_color() {
        use crate::ir::Color;
        let page = FixedPage {
            size: PageSize::default(),
            elements: vec![],
            background_color: Some(Color::new(255, 0, 0)),
        };
        assert_eq!(page.background_color, Some(Color::new(255, 0, 0)));
    }

    #[test]
    fn test_fixed_page_no_background_color() {
        let page = FixedPage {
            size: PageSize::default(),
            elements: vec![],
            background_color: None,
        };
        assert!(page.background_color.is_none());
    }
}
