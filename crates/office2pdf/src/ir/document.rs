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
    /// DOCX: a section that begins at the current page position instead of
    /// forcing a new page (`w:sectPr/w:type="continuous"`).
    FlowContinuous(FlowPage),
    /// PPTX: fixed coordinate pages.
    Fixed(FixedPage),
    /// XLSX: spreadsheet sheet pages.
    Sheet(SheetPage),
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
    /// The header this section's **first** page takes, where `<w:titlePg/>`
    /// asks for one. `None` means every page takes [`FlowPage::header`]
    /// (issue #846).
    pub first_header: Option<super::elements::HeaderFooter>,
    /// The footer this section's first page takes, under the same rule.
    pub first_footer: Option<super::elements::HeaderFooter>,
    /// Optional multi-column layout for the page.
    pub columns: Option<ColumnLayout>,
    /// Word document-grid line pitch in points (`w:docGrid w:linePitch`),
    /// present whenever the section carries a `w:docGrid` at all. That bare
    /// presence marks the file as authored in an East Asian Word edition,
    /// which is what decides the default tab stop (issue #393).
    pub line_grid_pitch: Option<f64>,
    /// Whether that grid snaps body lines to the pitch. Only a `w:docGrid
    /// w:type` of `lines`, `linesAndChars`, or `snapToChars` does; the
    /// `default` type an omitted attribute implies declares a pitch that Word
    /// then ignores for layout (issue #518).
    pub line_grid_snaps_lines: bool,
    /// Section page numbering (`w:sectPr/w:pgNumType`): where the counter
    /// restarts, and which numerals a `PAGE` field renders. `None` when the
    /// section declares nothing and simply continues.
    pub page_numbering: Option<PageNumbering>,
}

/// A section's `w:pgNumType`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageNumbering {
    /// `w:start`: the number this section's first page takes. `None` continues
    /// from the previous section.
    pub start: Option<u32>,
    /// `w:fmt`: the numerals a `PAGE` field renders in.
    pub format: PageNumberFormat,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PageNumberFormat {
    #[default]
    Decimal,
    LowerRoman,
    UpperRoman,
    LowerLetter,
    UpperLetter,
}

impl PageNumberFormat {
    /// The Typst numbering pattern that renders this format.
    pub fn typst_pattern(self) -> &'static str {
        match self {
            PageNumberFormat::Decimal => "1",
            PageNumberFormat::LowerRoman => "i",
            PageNumberFormat::UpperRoman => "I",
            PageNumberFormat::LowerLetter => "a",
            PageNumberFormat::UpperLetter => "A",
        }
    }
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
    /// Boxed: `Chart` is much the largest variant, and carrying it inline
    /// made every `FixedElement` pay for it (clippy's
    /// `large_enum_variant`).
    Chart(Box<super::elements::Chart>),
}

/// A spreadsheet sheet page (XLSX sheets).
#[derive(Debug, Clone)]
pub struct SheetPage {
    pub name: String,
    pub size: PageSize,
    pub margins: Margins,
    pub table: super::elements::Table,
    pub header: Option<super::elements::HeaderFooter>,
    pub footer: Option<super::elements::HeaderFooter>,
    /// Charts drawn on this sheet.
    pub charts: Vec<SheetChart>,
    /// Drawing images anchored within this sheet.
    pub images: Vec<SheetImage>,
    /// Drawing text boxes anchored within this sheet.
    pub text_boxes: Vec<SheetTextBox>,
    /// Drawing shapes with no text of their own anchored within this sheet.
    pub shapes: Vec<SheetShape>,
}

/// A worksheet drawing shape anchored to a sheet row: a DrawingML
/// `xdr:cxnSp` connector or a text-less `xdr:sp`, such as the line
/// separators Excel draws between charts (issue #1566).
#[derive(Debug, Clone)]
pub struct SheetShape {
    /// 1-indexed anchor row. Used only to order drawings deterministically;
    /// placement comes from `x_offset_pt`/`y_offset_pt` (issue #474).
    pub anchor_row: u32,
    /// Horizontal offset of the anchor from the sheet's left edge, points.
    pub x_offset_pt: f64,
    /// Vertical offset of the anchor from the sheet's content top, points.
    pub y_offset_pt: f64,
    /// Anchor width in points; a vertical line's is zero.
    pub width: f64,
    /// Anchor height in points; a horizontal line's is zero.
    pub height: f64,
    /// The geometry and paint, with line coordinates in points relative to
    /// the anchor's top-left corner.
    pub shape: super::elements::Shape,
}

/// A chart drawn on a worksheet.
#[derive(Debug, Clone)]
pub struct SheetChart {
    /// 1-indexed anchor row. Orders drawings deterministically; `u32::MAX`
    /// marks a chart that no drawing anchors.
    pub anchor_row: u32,
    /// Where the drawing anchor puts the chart, in points from the sheet's
    /// content origin. `None` for a chart no drawing anchors, which flows
    /// after the grid instead.
    pub placement: Option<SheetChartPlacement>,
    pub chart: super::elements::Chart,
}

/// The absolute placement a worksheet drawing anchor gives a chart.
///
/// Excel floats an anchored chart over the cells at worksheet coordinates and
/// sizes it to the anchor, exactly as it does a picture or a text box (issues
/// #459, #474). Flowing charts between row segments instead drew the reported
/// workbook's chart above the grid at an intrinsic size, leaving the band its
/// anchor reserves empty (issue #982).
#[derive(Debug, Clone, Copy)]
pub struct SheetChartPlacement {
    /// Horizontal offset of the anchor from the sheet's left edge, points.
    pub x_offset_pt: f64,
    /// Vertical offset of the anchor from the sheet's content top, points.
    pub y_offset_pt: f64,
    /// Width in points, from the columns the anchor spans, before
    /// `print_scale`.
    pub width: f64,
    /// Height in points, from the rows the anchor spans, before `print_scale`.
    pub height: f64,
    /// The fit-to-page scale the sheet prints at, applied to the drawing whole.
    /// `1.0` for a sheet that prints at full size.
    ///
    /// Excel scales a printed sheet whole, drawings included, so the chart's
    /// own text, tick marks and legend come down by the same factor as its
    /// frame. Shrinking the frame alone left the reported workbook's tick
    /// labels and legend entries at the size the chart XML declares, about 22%
    /// larger than Excel prints them (issue #1069).
    ///
    /// The chart lays itself out in `width` x `height` and the whole result is
    /// then scaled, so the box it occupies on the page is `width * print_scale`
    /// by `height * print_scale`. Scaling the chart's declared type sizes
    /// instead would not do: an axis title's size never reaches the IR, and
    /// several chart layout paths switch model on whether a size was declared
    /// at all, so filling one in changes the chrome even at a scale of 1.
    pub print_scale: f64,
    /// The page-column window the chart is clipped to, in points from the
    /// page's content left edge, after `print_scale`. Set by width
    /// pagination: Excel clips a chart at the printed edge of the tile it
    /// crosses and paints the remainder on the next tile, where the copy
    /// carries a negative `x_offset_pt` (issue #1598). `None` draws the chart
    /// unclipped, as a sheet that fits one page-column does.
    pub clip_window: Option<SheetClipWindow>,
}

/// The horizontal window of one printed page-column, in points from the
/// page's content left edge.
///
/// A window starts after any repeated print-title columns, so a drawing
/// continued from an earlier tile never paints over them.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SheetClipWindow {
    /// Where the tile's own columns begin.
    pub left_pt: f64,
    /// The tile's printed width.
    pub width_pt: f64,
}

/// A worksheet text box anchored to a sheet row.
#[derive(Debug, Clone)]
pub struct SheetTextBox {
    /// 1-indexed anchor row. Used only to order drawings deterministically;
    /// placement comes from `x_offset_pt`/`y_offset_pt` (issue #474).
    pub anchor_row: u32,
    /// Horizontal offset of the anchor from the sheet's left edge, points.
    pub x_offset_pt: f64,
    /// Vertical offset of the anchor from the sheet's content top, points
    /// (issue #474).
    pub y_offset_pt: f64,
    pub width: f64,
    pub height: f64,
    pub paragraphs: Vec<super::elements::Paragraph>,
    /// Box fill color.
    pub fill: Option<super::style::Color>,
    /// Box outline.
    pub border: Option<super::elements::BorderSide>,
    /// bodyPr anchor="ctr": center text vertically inside the box.
    pub vertical_center: bool,
}

/// A worksheet drawing image anchored to a sheet row.
#[derive(Debug, Clone)]
pub struct SheetImage {
    /// 1-indexed anchor row. Used only to order drawings deterministically;
    /// placement comes from `x_offset_pt`/`y_offset_pt` (issue #474).
    pub anchor_row: u32,
    /// Horizontal offset of the anchor from the sheet's left edge, points.
    pub x_offset_pt: f64,
    /// Vertical offset of the anchor from the sheet's content top, points.
    /// Excel overlays drawings on the grid at absolute worksheet
    /// coordinates rather than placing them between rows (issue #474).
    pub y_offset_pt: f64,
    pub image: super::elements::ImageData,
    /// Width of the page-column window this image is clipped to, from the
    /// page's content left edge. Set by drawing-width pagination: Excel clips
    /// a drawing at the printable edge and continues it on the next
    /// page-column, so a paged copy may also carry a negative `x_offset_pt`
    /// (issue #713). `None` draws the image unclipped — and so does `Some`
    /// when the image lacks a known width or height, since the renderer's
    /// clip box needs the image's own size.
    pub clip_width_pt: Option<f64>,
}

#[cfg(test)]
#[path = "document_tests.rs"]
mod tests;
