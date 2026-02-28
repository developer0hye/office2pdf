use super::style::{Color, ParagraphStyle, TextStyle};

/// Header or footer content for flow pages.
#[derive(Debug, Clone)]
pub struct HeaderFooter {
    pub paragraphs: Vec<HeaderFooterParagraph>,
}

/// A paragraph within a header or footer.
#[derive(Debug, Clone)]
pub struct HeaderFooterParagraph {
    pub style: ParagraphStyle,
    pub elements: Vec<HFInline>,
}

/// An inline element within a header or footer paragraph.
#[derive(Debug, Clone)]
pub enum HFInline {
    /// A text run with styling.
    Run(Run),
    /// Current page number field.
    PageNumber,
    /// Total page count field.
    TotalPages,
}

/// Block-level content elements.
#[derive(Debug, Clone)]
pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
    Image(ImageData),
    FloatingImage(FloatingImage),
    List(List),
    MathEquation(MathEquation),
    Chart(Chart),
    PageBreak,
}

/// A chart extracted from an embedded chart object.
#[derive(Debug, Clone)]
pub struct Chart {
    /// The type of chart (bar, line, pie, etc.).
    pub chart_type: ChartType,
    /// Optional chart title.
    pub title: Option<String>,
    /// Category labels (x-axis or pie slice names).
    pub categories: Vec<String>,
    /// Data series.
    pub series: Vec<ChartSeries>,
}

/// The type of chart.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartType {
    Bar,
    Column,
    Line,
    Pie,
    Area,
    Scatter,
    Other(String),
}

/// A data series within a chart.
#[derive(Debug, Clone)]
pub struct ChartSeries {
    /// Optional series name.
    pub name: Option<String>,
    /// Data values for this series.
    pub values: Vec<f64>,
}

/// A math equation (from OMML or similar).
#[derive(Debug, Clone)]
pub struct MathEquation {
    /// Typst math notation content (without surrounding `$` delimiters).
    pub content: String,
    /// Whether this is a display equation (centered, on its own line) vs inline.
    pub display: bool,
}

/// How text wraps around a floating image.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WrapMode {
    /// Text wraps around the image on both sides (square bounding box).
    Square,
    /// Text wraps tightly around the image contour.
    Tight,
    /// Text appears above and below the image only (no side wrapping).
    TopAndBottom,
    /// Image is behind the text (no wrapping, text flows over).
    Behind,
    /// Image is in front of the text (no wrapping, image covers text).
    InFront,
    /// No text wrapping.
    None,
}

/// A floating image with positioning and text wrap mode.
#[derive(Debug, Clone)]
pub struct FloatingImage {
    pub image: ImageData,
    pub wrap_mode: WrapMode,
    /// Horizontal offset in points from the anchor reference.
    pub offset_x: f64,
    /// Vertical offset in points from the anchor reference.
    pub offset_y: f64,
}

/// The kind of list: ordered (numbered) or unordered (bulleted).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListKind {
    Ordered,
    Unordered,
}

/// A list block containing items at various indent levels.
#[derive(Debug, Clone)]
pub struct List {
    pub kind: ListKind,
    pub items: Vec<ListItem>,
}

/// A single list item with content and indent level.
#[derive(Debug, Clone)]
pub struct ListItem {
    pub content: Vec<Paragraph>,
    pub level: u32,
}

/// A paragraph consisting of styled text runs.
#[derive(Debug, Clone)]
pub struct Paragraph {
    pub style: ParagraphStyle,
    pub runs: Vec<Run>,
}

/// A run of text with uniform formatting.
#[derive(Debug, Clone)]
pub struct Run {
    pub text: String,
    pub style: TextStyle,
    /// Optional hyperlink URL. When present, the run is rendered as a clickable link.
    pub href: Option<String>,
    /// Optional footnote/endnote content. When present, a footnote marker is emitted and
    /// the content is rendered at the bottom of the page.
    pub footnote: Option<String>,
}

/// A table.
#[derive(Debug, Clone)]
pub struct Table {
    pub rows: Vec<TableRow>,
    pub column_widths: Vec<f64>,
}

/// A table row.
#[derive(Debug, Clone)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
    pub height: Option<f64>,
}

/// A data bar rendering within a cell (conditional formatting).
#[derive(Debug, Clone)]
pub struct DataBarInfo {
    /// Bar color.
    pub color: Color,
    /// Fill percentage from 0.0 to 1.0.
    pub fill_pct: f64,
}

/// A table cell.
#[derive(Debug, Clone)]
pub struct TableCell {
    pub content: Vec<Block>,
    pub col_span: u32,
    pub row_span: u32,
    pub border: Option<CellBorder>,
    pub background: Option<Color>,
    /// DataBar conditional formatting render info.
    pub data_bar: Option<DataBarInfo>,
    /// IconSet text symbol prepended to cell content.
    pub icon_text: Option<String>,
}

impl Default for TableCell {
    fn default() -> Self {
        Self {
            content: Vec::new(),
            col_span: 1,
            row_span: 1,
            border: None,
            background: None,
            data_bar: None,
            icon_text: None,
        }
    }
}

/// Cell border specification.
#[derive(Debug, Clone)]
pub struct CellBorder {
    pub top: Option<BorderSide>,
    pub bottom: Option<BorderSide>,
    pub left: Option<BorderSide>,
    pub right: Option<BorderSide>,
}

/// A single border side.
#[derive(Debug, Clone)]
pub struct BorderSide {
    pub width: f64,
    pub color: Color,
}

/// Image data.
#[derive(Debug, Clone)]
pub struct ImageData {
    pub data: Vec<u8>,
    pub format: ImageFormat,
    pub width: Option<f64>,
    pub height: Option<f64>,
}

/// Supported image formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageFormat {
    Png,
    Jpeg,
    Gif,
    Bmp,
    Tiff,
}

impl ImageFormat {
    /// Return the file extension for this image format.
    pub fn extension(&self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpeg",
            Self::Gif => "gif",
            Self::Bmp => "bmp",
            Self::Tiff => "tiff",
        }
    }
}

/// SmartArt diagram content extracted from a presentation.
///
/// Contains text items extracted from the SmartArt data model.
/// Rendered as a simplified list or boxed layout since full SmartArt
/// layout engines are not feasible in a pure-Rust converter.
#[derive(Debug, Clone)]
pub struct SmartArt {
    /// Text items extracted from SmartArt data points (type="node").
    pub items: Vec<String>,
}

/// A single stop in a gradient fill.
#[derive(Debug, Clone)]
pub struct GradientStop {
    /// Position along the gradient axis, from 0.0 (start) to 1.0 (end).
    pub position: f64,
    /// Color at this stop.
    pub color: Color,
}

/// A linear gradient fill.
#[derive(Debug, Clone)]
pub struct GradientFill {
    /// Gradient color stops, ordered by position.
    pub stops: Vec<GradientStop>,
    /// Angle of the linear gradient in degrees (0 = left-to-right, 90 = top-to-bottom).
    pub angle: f64,
}

/// An outer shadow effect on a shape.
#[derive(Debug, Clone)]
pub struct Shadow {
    /// Blur radius in points.
    pub blur_radius: f64,
    /// Distance from the shape in points.
    pub distance: f64,
    /// Direction angle in degrees (0 = right, 90 = down, 180 = left, 270 = up).
    pub direction: f64,
    /// Shadow color.
    pub color: Color,
    /// Opacity from 0.0 (fully transparent) to 1.0 (fully opaque).
    pub opacity: f64,
}

/// Basic geometric shape.
#[derive(Debug, Clone)]
pub struct Shape {
    pub kind: ShapeKind,
    pub fill: Option<Color>,
    /// Gradient fill for the shape (takes precedence over solid fill when present).
    pub gradient_fill: Option<GradientFill>,
    pub stroke: Option<BorderSide>,
    /// Rotation angle in degrees (clockwise).
    pub rotation_deg: Option<f64>,
    /// Opacity from 0.0 (fully transparent) to 1.0 (fully opaque).
    pub opacity: Option<f64>,
    /// Outer shadow effect.
    pub shadow: Option<Shadow>,
}

/// Shape types.
#[derive(Debug, Clone)]
pub enum ShapeKind {
    Rectangle,
    Ellipse,
    Line {
        x2: f64,
        y2: f64,
    },
    /// Rectangle with rounded corners. `radius_fraction` is relative to `min(width, height)`.
    RoundedRectangle {
        radius_fraction: f64,
    },
    /// Arbitrary polygon defined by vertices normalized to 0.0–1.0 relative to the bounding box.
    Polygon {
        vertices: Vec<(f64, f64)>,
    },
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_table_cell_default() {
        let cell = TableCell::default();
        assert_eq!(cell.col_span, 1);
        assert_eq!(cell.row_span, 1);
        assert!(cell.content.is_empty());
        assert!(cell.border.is_none());
        assert!(cell.background.is_none());
    }

    #[test]
    fn test_list_item_default() {
        let item = ListItem {
            content: vec![Paragraph {
                style: ParagraphStyle::default(),
                runs: vec![Run {
                    text: "Item 1".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                }],
            }],
            level: 0,
        };
        assert_eq!(item.level, 0);
        assert_eq!(item.content.len(), 1);
    }

    #[test]
    fn test_list_unordered() {
        let list = List {
            kind: ListKind::Unordered,
            items: vec![
                ListItem {
                    content: vec![Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Bullet 1".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }],
                    level: 0,
                },
                ListItem {
                    content: vec![Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Bullet 2".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }],
                    level: 0,
                },
            ],
        };
        assert_eq!(list.kind, ListKind::Unordered);
        assert_eq!(list.items.len(), 2);
    }

    #[test]
    fn test_list_ordered() {
        let list = List {
            kind: ListKind::Ordered,
            items: vec![ListItem {
                content: vec![Paragraph {
                    style: ParagraphStyle::default(),
                    runs: vec![Run {
                        text: "Step 1".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }],
                }],
                level: 0,
            }],
        };
        assert_eq!(list.kind, ListKind::Ordered);
        assert_eq!(list.items.len(), 1);
    }

    #[test]
    fn test_list_nested() {
        let list = List {
            kind: ListKind::Unordered,
            items: vec![
                ListItem {
                    content: vec![Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Top".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }],
                    level: 0,
                },
                ListItem {
                    content: vec![Paragraph {
                        style: ParagraphStyle::default(),
                        runs: vec![Run {
                            text: "Nested".to_string(),
                            style: TextStyle::default(),
                            href: None,
                            footnote: None,
                        }],
                    }],
                    level: 1,
                },
            ],
        };
        assert_eq!(list.items[0].level, 0);
        assert_eq!(list.items[1].level, 1);
    }

    #[test]
    fn test_paragraph_with_runs() {
        let para = Paragraph {
            style: ParagraphStyle::default(),
            runs: vec![
                Run {
                    text: "Hello ".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                },
                Run {
                    text: "world".to_string(),
                    style: TextStyle {
                        bold: Some(true),
                        ..TextStyle::default()
                    },
                    href: None,
                    footnote: None,
                },
            ],
        };
        assert_eq!(para.runs.len(), 2);
        assert_eq!(para.runs[0].text, "Hello ");
        assert_eq!(para.runs[1].style.bold, Some(true));
    }

    #[test]
    fn test_header_footer_with_text() {
        let hf = HeaderFooter {
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![HFInline::Run(Run {
                    text: "My Header".to_string(),
                    style: TextStyle::default(),
                    href: None,
                    footnote: None,
                })],
            }],
        };
        assert_eq!(hf.paragraphs.len(), 1);
        assert_eq!(hf.paragraphs[0].elements.len(), 1);
        match &hf.paragraphs[0].elements[0] {
            HFInline::Run(r) => assert_eq!(r.text, "My Header"),
            _ => panic!("Expected Run"),
        }
    }

    #[test]
    fn test_header_footer_with_page_number() {
        let hf = HeaderFooter {
            paragraphs: vec![HeaderFooterParagraph {
                style: ParagraphStyle::default(),
                elements: vec![
                    HFInline::Run(Run {
                        text: "Page ".to_string(),
                        style: TextStyle::default(),
                        href: None,
                        footnote: None,
                    }),
                    HFInline::PageNumber,
                ],
            }],
        };
        assert_eq!(hf.paragraphs[0].elements.len(), 2);
        assert!(matches!(hf.paragraphs[0].elements[1], HFInline::PageNumber));
    }
}
