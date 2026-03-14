use std::collections::BTreeMap;

use super::style::{Alignment, Color, ParagraphStyle, TextStyle};

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
    FloatingTextBox(FloatingTextBox),
    List(List),
    MathEquation(MathEquation),
    Chart(Chart),
    PageBreak,
    ColumnBreak,
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

/// A floating text box with positioning, size, and text wrap mode.
#[derive(Debug, Clone)]
pub struct FloatingTextBox {
    pub content: Vec<Block>,
    pub wrap_mode: WrapMode,
    pub width: f64,
    pub height: f64,
    /// Horizontal offset in points from the anchor reference.
    pub offset_x: f64,
    /// Vertical offset in points from the anchor reference.
    pub offset_y: f64,
}

/// Vertical alignment for fixed text box content.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TextBoxVerticalAlign {
    #[default]
    Top,
    Center,
    Bottom,
}

/// A fixed-position text box with content padding and vertical alignment.
#[derive(Debug, Clone)]
pub struct TextBoxData {
    pub content: Vec<Block>,
    pub padding: Insets,
    pub vertical_align: TextBoxVerticalAlign,
    /// Background fill color for the text box.
    pub fill: Option<Color>,
    /// Opacity from 0.0 (fully transparent) to 1.0 (fully opaque).
    pub opacity: Option<f64>,
    /// Border stroke for the text box.
    pub stroke: Option<BorderSide>,
    /// Shape geometry when the text box originates from a non-rectangular shape
    /// (e.g., `roundRect`, `homePlate`). `None` means default rectangle.
    pub shape_kind: Option<ShapeKind>,
    /// When true, text should not wrap — the content width is unconstrained.
    /// Corresponds to `<a:bodyPr wrap="none"/>` in OOXML.
    pub no_wrap: bool,
}

/// The kind of list: ordered (numbered) or unordered (bulleted).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListKind {
    Ordered,
    Unordered,
}

/// Numbering configuration for a specific list level.
#[derive(Debug, Clone, PartialEq)]
pub struct ListLevelStyle {
    pub kind: ListKind,
    /// Optional Typst numbering pattern derived from Word's lvlText/numFmt.
    pub numbering_pattern: Option<String>,
    /// Whether parent numbers should be shown for nested ordered lists.
    pub full_numbering: bool,
    /// Optional concrete marker text for unordered PPTX bullet lists.
    pub marker_text: Option<String>,
    /// Optional concrete marker presentation resolved from the source format.
    pub marker_style: Option<TextStyle>,
}

/// A list block containing items at various indent levels.
#[derive(Debug, Clone)]
pub struct List {
    pub kind: ListKind,
    pub items: Vec<ListItem>,
    /// Per-level list style overrides. Levels not present fall back to `kind`.
    pub level_styles: BTreeMap<u32, ListLevelStyle>,
}

/// A single list item with content and indent level.
#[derive(Debug, Clone)]
pub struct ListItem {
    pub content: Vec<Paragraph>,
    pub level: u32,
    /// Ordered list item number when this item begins a new numbering run.
    pub start_at: Option<u32>,
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
#[derive(Debug, Clone, Default)]
pub struct Table {
    pub rows: Vec<TableRow>,
    pub column_widths: Vec<f64>,
    /// Number of leading rows that should repeat as the table header.
    pub header_row_count: usize,
    /// Optional block alignment for the table within the flow.
    pub alignment: Option<Alignment>,
    /// Default cell padding applied by the table when cells don't override it.
    pub default_cell_padding: Option<Insets>,
    /// When true, row heights should be derived from content instead of forced to
    /// the exact source row sizes. PowerPoint often renders slide tables this way.
    pub use_content_driven_row_heights: bool,
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

/// Vertical alignment within a table cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellVerticalAlign {
    Top,
    Center,
    Bottom,
}

/// Insets/padding in points.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Insets {
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
    pub left: f64,
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
    /// Vertical alignment of cell content.
    pub vertical_align: Option<CellVerticalAlign>,
    /// Optional cell padding override in points.
    pub padding: Option<Insets>,
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
            vertical_align: None,
            padding: None,
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

/// Border line style (dash pattern).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum BorderLineStyle {
    #[default]
    Solid,
    Dashed,
    Dotted,
    DashDot,
    DashDotDot,
    Double,
    None,
}

/// A single border side.
#[derive(Debug, Clone)]
pub struct BorderSide {
    pub width: f64,
    pub color: Color,
    pub style: BorderLineStyle,
}

/// Fractions of the source image cropped away from each edge.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct ImageCrop {
    pub left: f64,
    pub top: f64,
    pub right: f64,
    pub bottom: f64,
}

impl ImageCrop {
    pub fn is_empty(&self) -> bool {
        self.left == 0.0 && self.top == 0.0 && self.right == 0.0 && self.bottom == 0.0
    }
}

/// Image data.
#[derive(Debug, Clone)]
pub struct ImageData {
    pub data: Vec<u8>,
    pub format: ImageFormat,
    pub width: Option<f64>,
    pub height: Option<f64>,
    pub crop: Option<ImageCrop>,
    /// Optional border stroke around the image.
    pub stroke: Option<BorderSide>,
}

/// Supported image formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageFormat {
    Png,
    Jpeg,
    Gif,
    Bmp,
    Tiff,
    Svg,
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
            Self::Svg => "svg",
        }
    }
}

/// A node in a SmartArt diagram with hierarchy depth.
#[derive(Debug, Clone, PartialEq)]
pub struct SmartArtNode {
    /// The text content of this node.
    pub text: String,
    /// Depth in the hierarchy (0 = top-level node).
    pub depth: usize,
}

/// SmartArt diagram content extracted from a presentation.
///
/// Contains nodes extracted from the SmartArt data model with hierarchy
/// information derived from the connection list.
/// Rendered as an indented tree or numbered steps since full SmartArt
/// layout engines are not feasible in a pure-Rust converter.
#[derive(Debug, Clone)]
pub struct SmartArt {
    /// Nodes extracted from SmartArt data points with hierarchy depth.
    pub items: Vec<SmartArtNode>,
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
    /// Straight line from `(x1,y1)` to `(x2,y2)` in points, relative to element's top-left.
    Line {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        head_end: ArrowHead,
        tail_end: ArrowHead,
    },
    /// Multi-segment polyline in points, relative to element's top-left.
    Polyline {
        points: Vec<(f64, f64)>,
        head_end: ArrowHead,
        tail_end: ArrowHead,
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

/// Arrowhead decoration on a line endpoint.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ArrowHead {
    #[default]
    None,
    Triangle,
}

#[cfg(test)]
#[path = "elements_tests.rs"]
mod tests;
