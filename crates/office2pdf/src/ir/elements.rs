use std::collections::BTreeMap;

use super::style::{Alignment, Color, ParagraphStyle, TabLeader, TextStyle};

/// Header or footer content for flow pages and worksheets.
#[derive(Debug, Clone)]
pub struct HeaderFooter {
    pub paragraphs: Vec<HeaderFooterParagraph>,
    /// Distance in points from the page edge, when the source states one.
    pub distance_from_edge: Option<f64>,
    /// Fit-to-page scale of a worksheet header/footer's horizontal coordinate
    /// box. `None` for flow-page stories and sheets whose story is not scaled.
    pub sheet_print_scale: Option<f64>,
    /// Anchored shapes the story draws, positioned against the page rather
    /// than laid out in the story's flow. A header's decorative banner is one
    /// of these and carries no text at all (issue #961).
    pub shapes: Vec<HeaderFooterShape>,
}

/// A shape a header or footer story anchors to the page.
#[derive(Debug, Clone)]
pub struct HeaderFooterShape {
    pub shape: Shape,
    pub frame: HeaderFooterFrame,
    /// On-page bounding-box size in points, from `wp:extent`.
    pub width: f64,
    pub height: f64,
    /// `<wp:anchor behindDoc="1">` — drawn under the page's own content
    /// instead of over it, which is where a decorative banner belongs.
    pub behind_text: bool,
}

/// A paragraph within a header or footer.
#[derive(Debug, Clone)]
pub struct HeaderFooterParagraph {
    pub style: ParagraphStyle,
    pub elements: Vec<HFInline>,
    pub border: Option<CellBorder>,
    /// `w:pBdr` per-side `w:space` offsets in points, which set the gap Word
    /// leaves between the paragraph text and each rule.
    pub border_space: Option<Insets>,
    pub frame: Option<HeaderFooterFrame>,
}

/// Page- or margin-relative positioning for a header/footer paragraph frame.
#[derive(Debug, Clone, PartialEq)]
pub struct HeaderFooterFrame {
    pub x: Option<f64>,
    pub y: Option<f64>,
    pub width: Option<f64>,
    pub height: Option<f64>,
    pub horizontal_anchor: FrameAnchor,
    pub vertical_anchor: FrameAnchor,
    /// `<wp:align>` in place of an offset, resolved against the anchor's
    /// reference frame at render time because the parser does not know the
    /// page size. `None` means the offset above states the position
    /// (issue #847).
    pub horizontal_align: Option<FrameAlign>,
    pub vertical_align: Option<FrameAlign>,
    /// The text box's own left/top padding in points, applied *after* the
    /// alignment above resolves — the alignment pins the box, the padding sits
    /// inside it. Zero for a `w:framePr` frame, which has no padding of its
    /// own (issue #847).
    pub inset_left: f64,
    pub inset_top: f64,
    /// When the shape seats its text at its own bottom edge (`<a:bodyPr
    /// anchor="b">`), the gap between that edge and the reference frame's
    /// bottom — so the block is placed upward from the page rather than
    /// downward from the box's top, which is the only way its height enters
    /// the position (issue #847).
    pub bottom_offset: Option<f64>,
    /// Whether the box's paragraph wraps inside it. `<a:bodyPr wrap="none">`
    /// says it does not: the line stays whole and hangs out of the text column,
    /// which the same element's `horzOverflow="overflow"` then permits
    /// (issue #967). True for everything else, including a `w:framePr` frame.
    pub wraps_text: bool,
}

/// Which edge of the reference frame a `<wp:align>` pins a shape to.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FrameAlign {
    Start,
    Center,
    End,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum FrameAnchor {
    Page,
    Margin,
    #[default]
    Text,
}

/// A position-relative tab (`w:ptab`) inside header/footer content.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PositionedTab {
    pub alignment: PositionedTabAlignment,
    pub relative_to: PositionedTabRelativeTo,
    pub leader: TabLeader,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PositionedTabAlignment {
    Center,
    #[default]
    Left,
    Right,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PositionedTabRelativeTo {
    Indent,
    #[default]
    Margin,
}

/// An inline element within a header or footer paragraph.
#[derive(Debug, Clone)]
pub enum HFInline {
    /// A text run with styling.
    Run(Run),
    /// An inline image embedded in the header or footer part.
    Image(ImageData),
    /// Current page number field, carrying the run properties of the `w:r`
    /// that holds it so the number matches the surrounding literals.
    PageNumber(TextStyle),
    /// Total page count field, styled like [`HFInline::PageNumber(TextStyle::default())`].
    TotalPages(TextStyle),
    /// Alignment tab positioned relative to the paragraph indent or page margin.
    PositionedTab(PositionedTab),
}

/// Block-level content elements.
#[derive(Debug, Clone)]
pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
    Image(ImageData),
    /// Consecutive inline images from one flow paragraph.
    InlineImages(Vec<ImageData>),
    FloatingImage(FloatingImage),
    FloatingTextBox(FloatingTextBox),
    FloatingShape(FloatingShape),
    List(List),
    MathEquation(MathEquation),
    /// Boxed for the same reason [`super::document::FixedElementKind::Chart`]
    /// is: `Chart` is much the largest variant, and carrying it inline made
    /// every `Block` of every flow page pay for it (clippy's
    /// `large_enum_variant`).
    Chart(Box<Chart>),
    /// A `TOC` field's result, computed at render time from the document's own
    /// headings or captions.
    TableOfContents(TableOfContents),
    /// A paragraph numbered by a `SEQ` field — a figure or table caption.
    ///
    /// It renders exactly like the paragraph it wraps; the wrapper exists so a
    /// `TOC \a` list can collect it (issue #576).
    Caption(Caption),
    PageBreak,
    ColumnBreak,
}

/// What a `TOC` field collects.
///
/// Word stores the entries it last computed inside the field. A generated
/// document leaves the field dirty and empty for Word to fill on open, so the
/// entries have to be computed rather than read (issue #576).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TableOfContents {
    /// `TOC \o "1-3"`: paragraphs whose style carries `w:outlineLvl`, to this
    /// depth.
    Headings { depth: u8 },
    /// `TOC \a "Figure"`: the captions counted by that `SEQ` identifier.
    Captions { identifier: String },
}

/// A caption paragraph and the `SEQ` identifier numbering it.
#[derive(Debug, Clone)]
pub struct Caption {
    /// The `SEQ` identifier — `Figure`, `Table` — whose list collects this.
    pub identifier: String,
    /// The text a `TOC \a` list shows: the caption without the label and the
    /// field's number, which Word leaves out of the list.
    pub entry_text: String,
    pub paragraph: Paragraph,
}

/// What `c:chartSpace/c:spPr/a:ln` asks for around the whole chart area.
///
/// The three cases are visually opposite and the corpus holds all of them, so
/// one unconditional default would put a border on charts that ask for none and
/// the wrong border on charts that ask for their own (#637):
///
/// - `xlsx/poi/WithChart.xlsx` declares no `a:ln` at all — [`Self::Default`].
/// - `xlsx/poi/123233_charts.xlsx` and `pptx/oxp_CU018-Chart-Cached-Data-41.pptx`
///   declare `<a:ln><a:noFill/></a:ln>` — [`Self::Suppressed`].
/// - `xlsx/office2pdf_repository_workbook.xlsx` declares a 9360 EMU `#d9d9d9`
///   line and `pptx/chart-picture-bg.pptx` a 28575 EMU accent one —
///   [`Self::Explicit`].
#[derive(Debug, Clone, PartialEq, Default)]
pub enum ChartAreaOutline {
    /// No `a:ln` at all, so the automatic outline applies — and what that is
    /// depends on [`Chart::host`]: Excel and Word draw a thin one, PowerPoint
    /// draws none (issue #823).
    #[default]
    Default,
    /// `<a:ln><a:noFill/></a:ln>` — the file asks for no outline.
    Suppressed,
    /// An explicit line. Either component falls back to the default when the
    /// file leaves it out, or names a colour this parser cannot resolve.
    Explicit {
        /// `a:ln/@w` converted from EMU.
        width_pt: Option<f64>,
        /// The line's `a:solidFill/a:srgbClr`.
        color: Option<Color>,
    },
}

/// What `c:chartSpace/c:spPr` asks Office to paint behind the whole chart.
///
/// The declaration belongs to the chart space, outside `c:chart`, so it covers
/// the title as well as the plot. Absence and `<a:noFill/>` stay distinct even
/// though both currently render transparently: absence leaves the host's
/// automatic behavior available, while `noFill` explicitly disables it
/// (issue #1217).
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ChartAreaFill {
    /// No top-level fill declaration in the chart area's `c:spPr`.
    #[default]
    Unspecified,
    /// A top-level `<a:noFill/>` explicitly makes the chart area transparent.
    Transparent,
    /// A top-level `<a:solidFill>` with a literal or theme-resolved colour.
    Solid(Color),
}

/// The Office application and surface a chart came from.
///
/// Excel and PowerPoint disagree about what "the automatic chart-area outline"
/// is: Excel draws one and PowerPoint draws none, so the same
/// [`ChartAreaOutline::Default`] has to resolve differently depending on where
/// the chart part was found (issue #823). Excel also gives an anchored
/// worksheet chart and a chartsheet different legend geometry even though
/// both chart parts live in the same package (issue #1315). The chart part
/// itself says nothing about either distinction; only the loader knows.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ChartHost {
    /// A chart hosted by an Excel worksheet, whether drawing-anchored or an
    /// orphaned chart part with no drawing placement.
    Spreadsheet,
    /// A chart occupying an Excel chartsheet.
    ///
    /// Excel gives chartsheet chrome a separate layout regime from an
    /// anchored worksheet chart even though both live in a workbook.
    SpreadsheetChartsheet,
    /// A chart on a slide.
    Presentation,
    /// A chart in a document. Its automatic outline is unmeasured, so it keeps
    /// the spreadsheet's, which is what every chart took before #823.
    #[default]
    WordProcessing,
}

/// The label the parser gives a `<c:radarChart>`.
///
/// The radar family has no `ChartType` variant of its own, so the parser and
/// the renderer agree through this one string rather than repeating the
/// literal (issue #679).
pub const RADAR_CHART_LABEL: &str = "Radar Chart";

/// A chart extracted from an embedded chart object.
#[derive(Debug, Clone)]
pub struct Chart {
    /// The type of chart (bar, line, pie, etc.).
    pub chart_type: ChartType,
    /// `<c:holeSize val>` for a doughnut, as a percentage of the outer radius.
    /// `None` for every other type (issue #679).
    pub hole_size_percent: Option<u32>,
    /// Optional chart title.
    pub title: Option<String>,
    /// Category labels (x-axis or pie slice names).
    pub categories: Vec<String>,
    /// Data series.
    pub series: Vec<ChartSeries>,
    /// How a category's series share one bar.
    pub grouping: ChartGrouping,
    /// Where the legend sits, from `<c:legendPos>`.
    pub legend_position: LegendPosition,
    /// Whether the chart declares a `<c:legend>` at all, and did not switch it
    /// off with `<c:delete val="1"/>`.
    ///
    /// Separate from `legend_position`, which falls back to a default for
    /// every chart and so cannot distinguish "no legend" from "legend on the
    /// right" (issue #762).
    pub has_legend: bool,
    /// Title of the category axis, from `<c:catAx><c:title>`.
    pub category_axis_title: Option<String>,
    /// Title of the value axis, from `<c:valAx><c:title>`. Office writes it
    /// rotated a quarter turn anticlockwise along the axis.
    pub value_axis_title: Option<String>,
    /// Where the category axis puts its major tick marks, from
    /// `<c:catAx><c:majorTickMark>`.
    pub category_axis_major_tick_mark: AxisTickMark,
    /// Where the value axis puts its major tick marks, from
    /// `<c:valAx><c:majorTickMark>`.
    pub value_axis_major_tick_mark: AxisTickMark,
    /// What `<c:catAx><c:spPr>` says about the category axis' line.
    pub category_axis_line: ChartLine,
    /// What `<c:valAx><c:spPr>` says about the value axis' line.
    pub value_axis_line: ChartLine,
    /// `<c:valAx><c:majorUnit>` — the tick interval the part states.
    /// `None` leaves the interval to the automatic scale (issue #882).
    pub value_axis_major_unit: Option<f64>,
    /// `<c:valAx><c:scaling><c:min>` — the value the axis starts at.
    ///
    /// `None` leaves that end to the automatic scale, which starts at zero
    /// unless the data reaches below it. A stated minimum is what puts zero
    /// inside the plot rather than on its edge, and the category axis is drawn
    /// on the value-zero line wherever that lands (issue #1184).
    pub value_axis_min: Option<f64>,
    /// `<c:valAx><c:scaling><c:max>` — the value the axis ends at. `None`
    /// leaves that end to the automatic scale (issue #1184).
    pub value_axis_max: Option<f64>,
    /// What `<c:majorGridlines><c:spPr>` says about the gridlines' line.
    pub major_gridline_line: ChartLine,
    /// Whether `<c:catAx><c:delete>` switched the category axis off.
    pub category_axis_deleted: bool,
    /// Whether `<c:valAx><c:delete>` switched the value axis off.
    ///
    /// Office keeps the rest of a switched-off axis' settings — a hidden axis
    /// usually still carries `<c:majorTickMark val="out"/>` — so the flag is
    /// what decides whether the axis is drawn, not the settings beside it.
    pub value_axis_deleted: bool,
    /// How the bars of one category share the band it gets, from
    /// `<c:barChart>`. Charts outside the bar family carry the defaults.
    pub bar_band_layout: BarBandLayout,
    /// `accent1`..`accent6` of the theme the chart's package declares, in that
    /// order, for series that state no fill of their own.
    ///
    /// Empty when the package has no theme, or names fewer than six accents,
    /// in which case the renderer keeps its built-in palette. That palette is
    /// the Office 2013+ one, so a file built on any other theme was recoloured
    /// by it (issue #670).
    pub theme_accent_colors: Vec<Color>,
    /// What the chart area's own outline should be, from
    /// `c:chartSpace/c:spPr/a:ln` (#637).
    pub chart_area_outline: ChartAreaOutline,
    /// What the chart area's own background should be, from the top-level fill
    /// inside `c:chartSpace/c:spPr` (#1217).
    pub chart_area_fill: ChartAreaFill,
    /// The face every string the chart draws is set in, from
    /// `c:chartSpace/c:txPr/a:p/a:pPr/a:defRPr/a:latin@typeface`.
    ///
    /// `None` when the chart names none, which is the common case: the face
    /// then comes from the package theme's minor font, since chart text is
    /// body text. The parser leaves a `+mn-lt`/`+mj-lt` token unresolved
    /// because the chart part names no theme of its own — the loader that knew
    /// which package this came from substitutes the face, exactly as it does
    /// for [`Chart::theme_accent_colors`] (issue #668).
    pub text_font_family: Option<String>,
    /// Which Office application and surface this chart came from, used for
    /// host-specific chart chrome (issues #823 and #1315).
    pub host: ChartHost,
    /// Run properties `c:chartSpace/c:txPr` declares, which govern every string
    /// the chart draws unless a more specific `c:txPr` overrides them.
    pub text_style: ChartTextStyle,
    /// What `c:title/c:txPr` declares for the chart-area title alone.
    ///
    /// Office writes the title's size, weight and colour here rather than on
    /// the chart space, so a title reading its size off `text_style` gets the
    /// wrong one — `tests/fixtures/xlsx/any_sheets.xlsx` states a bare
    /// `<a:defRPr/>` for the chart space and `sz="1400" b="0"` over a #595959
    /// fill for the title (issue #1215).
    pub title_text_style: ChartTextStyle,
    /// What `c:legend/c:txPr` declares for every legend entry.
    ///
    /// The same fixture gives its legend a 9pt regular #595959 run while its
    /// chart space states none of those properties, so the legend cannot be
    /// represented by [`Chart::text_style`] alone (issue #1236).
    pub legend_text_style: ChartTextStyle,
    /// What `c:catAx/c:txPr` declares for the category labels alone.
    pub category_axis_text_style: ChartTextStyle,
    /// What `c:valAx/c:txPr` declares for the value tick labels alone.
    pub value_axis_text_style: ChartTextStyle,
    /// `<c:valAx><c:numFmt formatCode>` — how the value axis prints its tick
    /// labels. Outranks a series' cache format for the axis (issue #865).
    pub value_axis_number_format: Option<String>,
    /// `<c:autoTitleDeleted val="1"/>` — the chart declines the automatic
    /// title Office would otherwise supply: its single series' name, or the
    /// placeholder printed when nothing names one (issues #883 and #1146).
    pub auto_title_deleted: bool,
    /// Whether `<c:title>` is present but names no text of its own — it
    /// carries no `<c:tx>`.
    ///
    /// The application then supplies the string: a lone named series lends its
    /// name, and any other chart gets the placeholder Office writes into a new
    /// chart. A part carrying no `<c:title>` at all, or one whose title names
    /// its own text, leaves this false (issue #1146).
    pub has_automatic_title: bool,
    /// Where `c:title/c:layout/c:manualLayout` anchors the title box inside
    /// the full chart area. The Presentation renderer consumes this; other
    /// hosts currently keep their automatic centred title band. `None` keeps
    /// that automatic placement everywhere (issue #1423). See
    /// [`ChartTitleLayout`].
    pub title_layout: Option<ChartTitleLayout>,
    /// Where `c:plotArea/c:layout/c:manualLayout` puts the inner plot
    /// rectangle inside the chart area. `None` for the automatic layout, which
    /// is every chart that states nothing and every layout this does not model
    /// (issue #1182). See [`ChartPlotAreaLayout`].
    pub plot_area_layout: Option<ChartPlotAreaLayout>,
    /// The shapes the drawing part behind `<c:userShapes>` lays over the
    /// chart, in the order that part states them (issue #1186).
    ///
    /// Empty for a chart naming no such part — which the chart XML alone
    /// cannot settle, since the relationship it names is resolved by whichever
    /// package holds it, exactly as [`Chart::theme_accent_colors`] is.
    pub user_shapes: Vec<ChartUserShape>,
}

/// The top-left title-box anchor stated by
/// `c:title/c:layout/c:manualLayout`, as fractions of the full chart area's
/// width and height.
///
/// PowerPoint writes `xMode="edge"` and `yMode="edge"` for an explicitly
/// positioned title. The Presentation renderer supports this edge-mode pair;
/// Spreadsheet renderers currently preserve their automatic title placement.
/// In native exports of the page-8 and page-11 charts from `GENERAL
/// SERVICES.pptx`, `chart_left + x * chart_width` and
/// `chart_top + y * chart_height` are the title box edges; the glyphs begin at
/// the box's standard text inset (issue #1423). Factor-mode values remain
/// automatic because they are offsets from the application's computed layout,
/// not chart-relative edges.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartTitleLayout {
    /// `c:x` — title-box left edge as a fraction of chart-area width.
    pub x: f64,
    /// `c:y` — title-box top edge as a fraction of chart-area height.
    pub y: f64,
}

/// The inner plot rectangle `c:plotArea/c:layout/c:manualLayout` states, as
/// fractions of the chart area's own width and height.
///
/// Measured against native Excel for Mac 16 exports of
/// `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`, whose four charts each
/// carry one of these. Rewriting a single value of the `january income:` bar
/// chart's layout and re-exporting moves the plot exactly as this reads it:
/// `c:x` 0.30222 -> 0.1 slid the bars 48.58pt left across a 240.23pt printed
/// chart area (0.20222 x 240.23 = 48.58), and `c:w` 0.64139 -> 0.4 scaled every
/// bar by 0.6236, which is 0.4/0.64139 and not the 0.288 a right-edge reading
/// of `c:w` predicts. `c:y` and `c:h` behave the same way down the frame.
///
/// So `x`/`y` are the plot's top-left corner and `w`/`h` are its size, both as
/// fractions of the chart area — not offsets from the automatic layout, which
/// is what the `factor` modes mean and what the parser declines to read.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartPlotAreaLayout {
    /// `c:x` — left edge, as a fraction of the chart area's width.
    pub x: f64,
    /// `c:y` — top edge, as a fraction of the chart area's height.
    pub y: f64,
    /// `c:w` — width, as a fraction of the chart area's width.
    pub width: f64,
    /// `c:h` — height, as a fraction of the chart area's height.
    pub height: f64,
}

/// A shape the drawing part behind `<c:userShapes>` places over a chart.
///
/// Charts carry their own drawing layer: `c:chartSpace/c:userShapes` names a
/// part full of `cdr:` anchors, each holding an ordinary DrawingML shape, and
/// Office draws them over the finished chart. The `CASH FLOW` caption printed
/// left of the cash-flow plot in `tests/fixtures/xlsx/issue_1181_fit_to_height.xlsx`
/// is one (issue #1186).
///
/// Both anchor kinds put the shape's corner at a fraction of the chart area,
/// and differ only in how they state its size — see [`ChartUserShapeExtent`].
/// The `<a:xfrm>` Office caches beside the anchor is not read: it records the
/// resolved rectangle from whenever the shape was last edited, and that
/// workbook's cache is 5.00pt wider than its anchor states — 96.93pt against
/// the 91.93pt of `0.11685 x 786.7`.
#[derive(Debug, Clone)]
pub struct ChartUserShape {
    /// `cdr:from` — the shape's top-left corner, as `(x, y)` fractions of the
    /// chart area's width and height.
    pub from: (f64, f64),
    /// How far the shape reaches from that corner.
    pub extent: ChartUserShapeExtent,
    /// The shape's text body, one entry per `<a:p>`.
    pub paragraphs: Vec<Paragraph>,
    /// `<a:bodyPr>`'s text insets in points, defaulting to the DrawingML pair
    /// (0.1in left and right, 0.05in top and bottom).
    pub text_insets: Insets,
    /// `<a:solidFill>` on the shape itself, if it states one.
    pub fill: Option<Color>,
    /// `<a:ln>` on the shape itself, if it states one.
    pub border: Option<BorderSide>,
    /// `<a:bodyPr wrap="none"/>` — the text runs on past the shape's own
    /// width instead of wrapping inside it, which is how Excel draws the
    /// reported caption.
    pub no_wrap: bool,
}

/// How a chart drawing's anchor states the size of the shape it holds.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ChartUserShapeExtent {
    /// `<cdr:relSizeAnchor>` states the opposite corner, as fractions of the
    /// chart area, so the shape scales with the chart.
    Corner { x: f64, y: f64 },
    /// `<cdr:absSizeAnchor>` states a size in EMU, held here in points, which
    /// stays put however the chart is resized.
    Size { width: f64, height: f64 },
}

/// Run properties a `c:txPr` declares for the strings it governs.
///
/// Every field is `Option` because "said nothing" and "said this" have to stay
/// distinguishable: a `c:catAx/c:txPr` that sets only `b` must still take its
/// size and character spacing from `c:chartSpace/c:txPr`, and an element that
/// declares no `c:txPr` at all must fall through to the renderer's default
/// (issues #669 and #1011).
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct ChartTextStyle {
    /// `a:defRPr@sz`, in points — the attribute is in hundredths.
    pub size_pt: Option<f64>,
    /// `a:defRPr@b`.
    pub bold: Option<bool>,
    /// `a:defRPr@spc`, in hundredths of a point.
    ///
    /// Keeping DrawingML's integer unit avoids inflating every `Chart` by
    /// three nullable `f64` values; conversion belongs at the rendering edge.
    pub letter_spacing_hundredths: Option<i32>,
    /// `a:defRPr/a:solidFill` — the colour the runs are set in (issue #916).
    pub color: Option<Color>,
    /// `a:bodyPr@vertOverflow="ellipsis"`. This body property is kept beside
    /// the run properties because every chart text scope already owns one
    /// `ChartTextStyle` (issue #1012).
    pub ellipsis_overflow: bool,
}

impl ChartTextStyle {
    /// This style's size where `override_style` states none.
    ///
    /// Office resolves a chart string against the most specific `c:txPr` that
    /// mentions the attribute, so an axis setting only `b` keeps the chart
    /// space's size rather than dropping to a default.
    pub fn resolved_size_pt(self, override_style: Self) -> Option<f64> {
        override_style.size_pt.or(self.size_pt)
    }

    /// This style's weight where `override_style` states none.
    pub fn resolved_bold(self, override_style: Self) -> Option<bool> {
        override_style.bold.or(self.bold)
    }

    /// This style's character spacing where `override_style` states none.
    pub fn resolved_letter_spacing(self, override_style: Self) -> Option<f64> {
        override_style
            .letter_spacing_hundredths
            .or(self.letter_spacing_hundredths)
            .map(|hundredths| hundredths as f64 / 100.0)
    }

    /// This style's colour where `override_style` states none.
    pub fn resolved_color(self, override_style: Self) -> Option<Color> {
        override_style.color.or(self.color)
    }
}

impl Chart {
    /// Every string the chart draws, concatenated.
    ///
    /// [`Chart::text_font_family`] is one face for all of them, so the fallback
    /// chain behind it has to cover each script that appears anywhere in the
    /// chart — a Korean category label needs the East Asian chain that a Latin
    /// family alone would not reach. Both the renderer, which emits the chain,
    /// and the font-context gate, which decides whether the search paths are
    /// scanned at all, have to sample the same strings (issue #668).
    pub fn text_sample(&self) -> String {
        self.title
            .iter()
            .chain(self.category_axis_title.iter())
            .chain(self.value_axis_title.iter())
            .chain(self.categories.iter())
            .chain(self.series.iter().filter_map(|series| series.name.as_ref()))
            .fold(String::new(), |mut sample, text| {
                sample.push_str(text);
                sample
            })
    }

    /// The family that plots `series` — the one it names, else the chart's own.
    ///
    /// Only a combo plot area names a family per series; see
    /// [`ChartSeries::plot_type`] (issue #1067).
    pub fn plot_type_of<'a>(&'a self, series: &'a ChartSeries) -> &'a ChartType {
        series.plot_type.as_ref().unwrap_or(&self.chart_type)
    }
}

/// How a bar chart's bars divide the band one category gets, from
/// `<c:barChart><c:gapWidth>` and `<c:barChart><c:overlap>`.
///
/// Both are measured in units of ONE bar's thickness rather than of the band,
/// so the two together decide how thick a bar is: a band holds the cluster its
/// series form plus a `gap_width_percent` gutter beside it.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct BarBandLayout {
    /// `<c:gapWidth>` (`ST_GapAmount`, 0..=500) — the gutter between
    /// neighbouring category bands, as a percentage of one bar's thickness.
    /// 100 makes the gutter exactly as wide as a bar.
    pub gap_width_percent: f64,
    /// `<c:overlap>` (`ST_Overlap`, -100..=100) — how far each clustered
    /// series' bar slides over its predecessor, as a percentage of one bar's
    /// thickness. Negative values push them apart instead.
    pub overlap_percent: f64,
}

impl Default for BarBandLayout {
    /// The values Office draws when a chart declares neither element, which are
    /// also ECMA-376's attribute defaults.
    ///
    /// Measured, not recalled: `tests/fixtures/xlsx/chart_sheet.xlsx` omits both
    /// elements, and Excel 16.0 exports its two clustered series as touching
    /// 42.3pt bars on a 148.1pt band — 148.1/42.3 is 2 series + 150%, and
    /// touching bars are an overlap of 0.
    fn default() -> Self {
        Self {
            gap_width_percent: 150.0,
            overlap_percent: 0.0,
        }
    }
}

/// Which side of an axis line its major tick marks project from, from
/// `<c:majorTickMark>` (`ST_TickMark`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum AxisTickMark {
    /// `none` — the axis carries no major tick marks.
    None,
    /// `in` — the ticks reach into the plot area.
    Inside,
    /// `out` — the ticks reach away from the plot area.
    ///
    /// The default is what Office renders for an axis that never mentions tick
    /// marks, not the `cross` ECMA-376 gives the attribute: Excel 16.0 exports
    /// `tests/fixtures/xlsx/WithChart.xlsx` — written by Apache POI without a
    /// single `<c:majorTickMark>` — with outward ticks on both axes.
    #[default]
    Outside,
    /// `cross` — the ticks straddle the axis line, reaching both ways.
    Cross,
}

/// Where a data label sits relative to the point it belongs to, from
/// `<c:dLblPos>` (ECMA-376 §21.2.2.49).
///
/// The element is optional and its default depends on the plot: a `clustered`
/// bar puts labels just beyond the bar's end, a stacked one centres them on
/// the segment, since an outside label would land on the segment above.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DataLabelPosition {
    /// `ctr` — centred on the point.
    #[default]
    Center,
    /// `outEnd` — just beyond the end of the bar or segment.
    OutsideEnd,
    /// `inEnd` — inside it, against the end.
    InsideEnd,
    /// `inBase` — inside it, against the baseline.
    InsideBase,
}

/// What a chart's data labels print, from `<c:dLbls>`.
///
/// Office joins the enabled parts with `<c:separator>`, defaulting to `"; "`.
///
/// Not `Eq`: [`Self::text_style`] carries a size in points.
#[derive(Debug, Clone, PartialEq)]
pub struct DataLabels {
    /// `<c:showVal>` — the point's own value.
    pub show_value: bool,
    /// `<c:showCatName>` — the category it sits over.
    pub show_category: bool,
    /// `<c:showSerName>` — the series it belongs to.
    pub show_series: bool,
    /// `<c:showPercent>` — its share of the category total.
    pub show_percent: bool,
    /// `<c:separator>` between the enabled parts.
    pub separator: String,
    /// `<c:dLbls><c:numFmt formatCode>` — how the label prints its value.
    /// Outranks the series' cache format, which is the source cell's own
    /// (issue #865).
    pub number_format: Option<String>,
    /// `<c:dLbls><c:txPr>` — the run properties the labels are set in. Its
    /// size outranks the chart space's, and where it states none the chart
    /// space's stands (issue #970).
    pub text_style: ChartTextStyle,
    /// `<c:dLblPos>`, or the default the plot's grouping implies when the
    /// part states none (issue #901).
    pub position: DataLabelPosition,
    /// Whether `<c:dLblPos>` was stated. A stated position outranks the
    /// grouping's default, so the two cannot be told apart by value alone —
    /// `ctr` is both a legal statement and the stacked default.
    pub position_stated: bool,
}

impl Default for DataLabels {
    fn default() -> Self {
        Self {
            show_value: false,
            show_category: false,
            show_series: false,
            show_percent: false,
            number_format: None,
            separator: "; ".to_string(),
            text_style: ChartTextStyle::default(),
            position: DataLabelPosition::Center,
            position_stated: false,
        }
    }
}

impl DataLabels {
    /// Whether anything at all is printed.
    pub fn is_empty(&self) -> bool {
        !(self.show_value || self.show_category || self.show_series || self.show_percent)
    }
}

/// Where a chart's legend sits relative to its plot, from `<c:legendPos>`.
///
/// ECMA-376 gives `ST_LegendPos` a default of `r`, which is also where every
/// legend used to be drawn.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum LegendPosition {
    Bottom,
    Left,
    #[default]
    Right,
    Top,
    TopRight,
}

impl LegendPosition {
    /// Whether entries flow left to right rather than stacking downward.
    /// PowerPoint lays a legend out along the edge it sits on.
    pub fn is_horizontal(self) -> bool {
        matches!(self, LegendPosition::Bottom | LegendPosition::Top)
    }
}

/// How the series of one category are combined, from `<c:grouping>`.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum ChartGrouping {
    /// Each series gets its own mark side by side, and so the shape a chart
    /// without `<c:grouping>` takes: ECMA-376 defaults `CT_BarGrouping` to
    /// `clustered` and `CT_Grouping` to `standard`, which both mean unstacked.
    #[default]
    Clustered,
    /// A category's series stack into one bar whose length is their total.
    Stacked,
    /// As `Stacked`, with every stack normalised to 100%.
    PercentStacked,
}

/// The type of chart.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ChartType {
    Bar,
    Column,
    Line,
    Pie,
    /// A pie with a concentric hole; `Chart::hole_size_percent` carries the
    /// inner radius as a percentage of the outer (issue #679).
    Doughnut,
    Area,
    Scatter,
    Other(String),
}

/// The point symbol a series' `<c:marker><c:symbol>` names.
///
/// Only the values this renderer can actually draw are listed; `dash`, `dot`,
/// `plus`, `star` and `picture` have no shape here, so a series naming one is
/// left to the automatic cycle rather than drawn as some other symbol.
/// ECMA-376 §21.2.3.29 `ST_MarkerStyle` spells all of them.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MarkerSymbol {
    /// `none` — the series draws no point markers at all.
    Off,
    Circle,
    Diamond,
    Square,
    Triangle,
    /// `x` — a diagonal cross.
    Cross,
}

/// A data series within a chart.
#[derive(Debug, Clone)]
pub struct ChartSeries {
    /// Optional series name.
    pub name: Option<String>,
    /// Data values for this series.
    pub values: Vec<f64>,
    /// Fill declared by the series' own `<c:spPr>`. `None` falls back to the
    /// built-in palette.
    pub fill: Option<Color>,
    /// Per-point fills from `<c:dPt>`, indexed by data point. A point's own
    /// fill outranks the series'; entries are `None` where the point declares
    /// none, and the vector may be shorter than `values`.
    pub point_fills: Vec<Option<Color>>,
    /// What this series' `<c:dLbls>` prints beside each point.
    pub data_labels: DataLabels,
    /// The number format code from `<c:numCache><c:formatCode>`, when the
    /// series states one that is not `General`. A ratio is stored as a
    /// fraction and only this says to print it as a percentage (issue #865).
    pub number_format: Option<String>,
    /// The plot-area family that declared this series, when it is not the
    /// chart's own [`Chart::chart_type`].
    ///
    /// A `c:plotArea` may hold one element per chart family, so a workbook can
    /// put `<c:barChart>` beside `<c:lineChart>` and expect stacked columns
    /// with a line over them. `chart_type` names only the family that governs
    /// the axis, so before this every series drew as that one kind and the
    /// columns disappeared into polylines (issue #1067).
    ///
    /// `None` — the ordinary case — means the chart's own family, which is
    /// what every series of a single-family chart is.
    pub plot_type: Option<ChartType>,
    /// The point symbol this series' own `<c:marker><c:symbol>` names.
    ///
    /// `None` means the file named none this renderer draws, so the automatic
    /// shape cycle picks one from the series index (issue #635). A file that
    /// does name one gets that symbol whatever its index, which is what issue
    /// #1107 was: a fourth series declaring `circle` drew the cycle's cross.
    pub marker_symbol: Option<MarkerSymbol>,
    /// Weight of the stroke this series is plotted with, from its own
    /// `<c:spPr><a:ln w="…"/>`, in points.
    ///
    /// `None` — no `<a:ln>`, or one stating only a colour — leaves the
    /// renderer's default weight. A stated width is Excel's: a workbook
    /// declaring `w="28440"` prints 2.24pt where the flat constant printed
    /// 2.0pt, thin enough to read beside gridlines that agree to the point
    /// (issue #1113).
    ///
    /// Only the families that plot a line read this — the line, radar and
    /// mixed-plot polylines and the legend key that samples them. A bar
    /// series' `<a:ln>` is its outline, which is a separate thing.
    pub line_width_pt: Option<f64>,
}

impl ChartSeries {
    /// The fill for one data point: its own, else the series', else `None` for
    /// the caller to take from the palette.
    pub fn fill_for_point(&self, point_index: usize) -> Option<Color> {
        self.point_fills
            .get(point_index)
            .copied()
            .flatten()
            .or(self.fill)
    }
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
    /// Clockwise rotation of the whole box about its centre, from the WPS
    /// shape's `<a:xfrm rot>`.
    pub shape_rotation_deg: Option<f64>,
    pub padding: Insets,
    pub vertical_align: TextBoxVerticalAlign,
    /// Horizontal offset in points from the anchor reference.
    pub offset_x: f64,
    /// Vertical offset in points from the anchor reference.
    pub offset_y: f64,
}

/// A floating geometric shape (rectangle, line/arrow, ellipse, …) positioned
/// with an anchor offset. Used for DrawingML word-processing shapes (`wps:wsp`)
/// that carry geometry but no text box — these have no docx-rs representation
/// and would otherwise be dropped (issue #176).
#[derive(Debug, Clone)]
pub struct FloatingShape {
    pub shape: Shape,
    /// On-page bounding-box width in points (from `wp:extent`).
    pub width: f64,
    /// On-page bounding-box height in points (from `wp:extent`).
    pub height: f64,
    /// Horizontal offset in points from the anchor reference.
    pub offset_x: f64,
    /// Vertical offset in points from the anchor reference.
    pub offset_y: f64,
    pub wrap_mode: WrapMode,
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
    /// Whether the renderer must dynamically shrink text to the box. This is
    /// set for `<a:normAutofit/>` only when PowerPoint did not save a
    /// `fontScale` or `lnSpcReduction`; saved results are applied while
    /// parsing. `<a:spAutoFit/>` grows the shape instead (issue #898).
    pub auto_fit: bool,
    /// Clockwise text rotation from `<a:bodyPr vert>` ("vert" = 90°,
    /// "vert270" = 270°); the box geometry itself stays unrotated.
    pub text_rotation_deg: Option<f64>,
    /// Clockwise rotation of the whole box about its centre: the shape's own
    /// `<a:xfrm rot>` composed with the angle of any rotated ancestor
    /// `<p:grpSp>`. Unlike `text_rotation_deg` the content lays out in the
    /// unrotated width x height box and the result is turned as a unit.
    pub shape_rotation_deg: Option<f64>,
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
    pub footnote: Option<Vec<Run>>,
}

/// A table.
#[derive(Debug, Clone, Default)]
pub struct Table {
    pub rows: Vec<TableRow>,
    pub column_widths: Vec<f64>,
    /// Number of leading rows that should repeat as the table header.
    pub header_row_count: usize,
    /// Number of rows above the repeating header that belong to the header
    /// block but must not repeat. Excel's `_xlnm.Print_Titles` can name a row
    /// below the sheet top; the rows above it print once, on the first page.
    pub non_repeating_header_row_count: usize,
    /// Optional block alignment for the table within the flow.
    pub alignment: Option<Alignment>,
    /// Default cell padding applied by the table when cells don't override it.
    pub default_cell_padding: Option<Insets>,
    /// When true, row heights should be derived from content instead of forced to
    /// the exact source row sizes. PowerPoint often renders slide tables this way.
    pub use_content_driven_row_heights: bool,
    /// Default vertical alignment for cells that don't override it.
    /// Excel prints cells bottom-aligned by default; Word/PowerPoint keep
    /// the renderer default (top).
    pub default_vertical_align: Option<CellVerticalAlign>,
    /// When true, a bottom-aligned cell rests its last line's descender on the
    /// row's bottom inset edge, as Excel prints. Only spreadsheet tables set
    /// this: Word's and PowerPoint's bottom-cell seating is unverified against
    /// native GT, so their emission must not change (issue #618).
    pub seats_bottom_aligned_text_on_descender: bool,
    /// The gap that descender seat never comes closer than to the row's bottom
    /// boundary, in points, however small the font (issues #1097, #1199). The
    /// independently measured floor matrix gives the remapped Calibri/Aptos
    /// family 3pt and the script-face theme family 4pt; later face-specific
    /// row-track mappings do not change that split. Zero for Word and
    /// PowerPoint tables, which take no descender seat at all.
    pub bottom_aligned_descent_floor_pt: f64,
    /// How borders are painted relative to the nominal grid boundary.
    pub border_paint_model: TableBorderPaintModel,
    /// When true, `<printOptions gridLines="1"/>` asks Excel to print its
    /// gridline hairline on every cell boundary of the printed range, under
    /// any explicit border styling (issue #622). Only spreadsheet tables set
    /// this, and it is honoured only together with
    /// `TableBorderPaintModel::ExcelBoundaryBands` machinery the
    /// gridlines reuse; Word/PowerPoint tables never print gridlines.
    pub prints_gridlines: bool,
    /// When true, `<printOptions headings="1"/>` prints Excel's row-number
    /// gutter and column-letter strip on every page (issue #623). The XLSX
    /// parser materializes both in the IR — the gutter as a prepended first
    /// column so the numbers flow with row pagination, and the letter strip
    /// as `rows[0]` — and codegen re-emits that first row as a
    /// `table.header(repeat: true)` above any print-title headers and paints
    /// GT's 1pt black print frame on the table's exterior boundaries.
    /// `header_row_count` and `non_repeating_header_row_count` keep counting
    /// from the first row AFTER the strip. Word/PowerPoint tables never set
    /// this.
    pub prints_headings: bool,
    /// When true, `<printOptions horizontalCentered="1"/>` centres the sheet's
    /// printed grid between the left and right print margins instead of
    /// printing it flush to the left one (issue #1110). Only spreadsheet
    /// tables set this; the columns it centres are the ones on this page, so
    /// a sheet split into column groups centres each group by its own width,
    /// as Excel prints it. Word/PowerPoint place a table box with
    /// [`Self::alignment`] instead.
    pub centers_between_print_margins: bool,
    /// The fit-to-page scale a sheet's sizes have **already** been multiplied
    /// by, where one applies (issues #1163, #1238). `None` on an unscaled
    /// sheet and on every Word/PowerPoint table.
    ///
    /// The parser folds the scale into every width, height and type size, so
    /// nothing downstream has to know it — except a rule Excel evaluates at
    /// the *declared* size and scales afterwards. Its wrapped-line advance is
    /// one: the sheet of issue #1163 prints its 14pt Segoe UI panel 17.22pt
    /// per line, which is the unscaled 21.00pt times this 0.82, not a whole
    /// number of points in the scaled domain. Its whole-point line seats and
    /// glyph advances follow the same order (issue #1238). Fitted sheet paint
    /// is another exception: Excel constructs the paper-space origin and its
    /// positive-axis background bleed before applying this scale (issue
    /// #1538).
    pub print_scale: Option<f64>,
}

/// How a table paints borders relative to its grid boundaries.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TableBorderPaintModel {
    /// Typst's native border stroke, centred on the boundary.
    #[default]
    CenteredStroke,
    /// Excel's printed bands, measured in #619 on the positive axis side for
    /// thin rules and with weight-specific offsets for wider rules.
    ExcelBoundaryBands,
    /// Word's filled border rectangles, anchored at the boundary and painted
    /// on its positive-axis side (issue #724).
    WordPositiveAxisBands,
}

/// A table row.
#[derive(Debug, Clone)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
    /// The row's fixed height in points, from a `w:trHeight` whose
    /// `@w:hRule` is `exact`.
    pub height: Option<f64>,
    /// The row's floor in points. Word supplies it through a `w:trHeight`
    /// whose `@w:hRule` is `atLeast` (the schema default); PowerPoint supplies
    /// the same growable constraint through `a:tr/@h`. The row is at least
    /// this tall and grows past it for taller content, which separates it from
    /// [`Self::height`] (issues #965 and #1253).
    pub minimum_height: Option<f64>,
}

/// Glyphs the parser records in [`TableCell::icon_text`] for Excel's arrow
/// icon sets. The renderer recognizes them to draw Excel's filled arrow shapes
/// instead of a character.
pub const ICON_ARROW_UP: &str = "\u{25B2}"; // ▲ black up-pointing triangle
pub const ICON_ARROW_DOWN: &str = "\u{25BC}"; // ▼ black down-pointing triangle
pub const ICON_ARROW_RIGHT: &str = "\u{25B6}"; // ▶ black right-pointing triangle
pub const ICON_ARROW_UP_RIGHT: &str = "\u{25E5}"; // ◥ black upper-right triangle
pub const ICON_ARROW_DOWN_RIGHT: &str = "\u{25E2}"; // ◢ black lower-right triangle

/// Glyph the parser records for the circular icon sets — traffic lights and
/// signs. Excel draws these as filled discs rather than characters, so the
/// renderer recognizes this marker the same way it does the arrows (#536).
pub const ICON_CIRCLE: &str = "\u{25CF}"; // ● black circle

/// The paint Excel gives one icon-set icon that it prints as a shaded sprite.
///
/// Excel's arrows are not flat shapes: the interior ramps between two colours
/// down the icon box's diagonal, and the silhouette carries a flat outline in a
/// saturated dark hue of its own that is *not* a darkening of the interior —
/// the amber band's is `#D87103` against an interior around `#FEE489`.
/// [`TableCell::icon_color`] can only carry one colour, so a band measured
/// against a native export carries this beside it (issue #1134).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct IconShading {
    /// Interior colour at the icon box's top-left corner.
    pub fill_start: Color,
    /// Interior colour at its bottom-right corner.
    pub fill_end: Color,
    /// The silhouette's outline, one flat colour across the whole sprite.
    pub outline: Color,
}

/// A data bar rendering within a cell (conditional formatting).
#[derive(Debug, Clone)]
pub struct DataBarInfo {
    /// Bar color.
    pub color: Color,
    /// Fill percentage in percentage points from 0.0 to 100.0.
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
    /// Opacity declared by the cell background color, from transparent `0.0`
    /// through opaque `1.0`. `None` keeps the legacy opaque paint.
    pub background_alpha: Option<f64>,
    /// DataBar conditional formatting render info.
    pub data_bar: Option<DataBarInfo>,
    /// IconSet text symbol prepended to cell content.
    pub icon_text: Option<String>,
    /// Fill color of the IconSet symbol (Excel draws icons in band colors).
    ///
    /// This is the band's flat stand-in, and the whole of its paint only for a
    /// band with no [`Self::icon_shading`].
    pub icon_color: Option<Color>,
    /// Excel's measured sprite paint for this band, where a native export has
    /// been read. `None` leaves the icon flat in [`Self::icon_color`].
    pub icon_shading: Option<IconShading>,
    /// Width in points that an unwrapped cell's single line paints across
    /// before it is clipped. `None` when the text fits its column and needs no
    /// clip box.
    ///
    /// Excel never moves a `wrapText="false"` cell's text to a second line, so
    /// this is what varies instead of the line count. A general/left cell paints
    /// on across consecutive empty columns to its right, giving its own column
    /// plus those; a centred or right-aligned cell, and any cell whose neighbour
    /// is occupied, gets its own column width alone and is clipped at its edge.
    pub spill_width: Option<f64>,
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
            background_alpha: None,
            data_bar: None,
            icon_text: None,
            icon_color: None,
            icon_shading: None,
            spill_width: None,
            vertical_align: None,
            padding: None,
        }
    }
}

/// Cell border specification.
#[derive(Debug, Clone, Default)]
pub struct CellBorder {
    pub top: Option<BorderSide>,
    pub bottom: Option<BorderSide>,
    pub left: Option<BorderSide>,
    pub right: Option<BorderSide>,
}

/// Border line style (dash pattern).
///
/// The first block is the cross-format set: Word `w:val` and Excel border
/// styles map onto it, and so do the three DrawingML presets that share its
/// names. The second block exists because DrawingML has more distinct dash
/// rhythms than that set can name — `lgDash` (8w on) and `sysDash` (3w on)
/// are not the same line as `dash` (4w on), and folding them together renders
/// one preset as another (issue #758). Word and Excel never produce them.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum BorderLineStyle {
    #[default]
    Solid,
    /// DrawingML `dash`.
    Dashed,
    /// DrawingML `dot`.
    Dotted,
    /// DrawingML `dashDot`.
    DashDot,
    /// No DrawingML preset maps here; Word `dotDotDash` and its Excel kin do.
    DashDotDot,
    Double,
    None,
    /// DrawingML `sysDot`.
    SystemDot,
    /// DrawingML `sysDash`.
    SystemDash,
    /// DrawingML `lgDash`.
    LargeDash,
    /// DrawingML `sysDashDot`.
    SystemDashDot,
    /// DrawingML `lgDashDot`.
    LargeDashDot,
    /// DrawingML `sysDashDotDot`.
    SystemDashDotDot,
    /// DrawingML `lgDashDotDot`.
    LargeDashDotDot,
}

/// How a stroke turns a corner.
///
/// DrawingML spells this as at most one of `a:round`, `a:bevel` or `a:miter`
/// inside `a:ln`; naming none of them selects `Round`, which is why that is the
/// default here (issue #1090).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LineJoin {
    #[default]
    Round,
    Bevel,
    Miter,
}

/// A single border side.
///
/// `join` describes a DrawingML `a:ln` and only shape and picture outlines
/// render it; Word and Excel have no corresponding border property, so their
/// sides leave it at the default and their codegen never writes it out.
#[derive(Debug, Clone)]
pub struct BorderSide {
    pub width: f64,
    pub color: Color,
    pub style: BorderLineStyle,
    pub join: LineJoin,
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
    /// Clockwise rotation in degrees about the image's centre: the picture's
    /// own `a:xfrm/@rot`, composed with the angle of any rotated ancestor
    /// `<p:grpSp>` in PPTX. `None` means upright (issues #682, #895, #1366).
    pub rotation_deg: Option<f64>,
    /// Mirror a fixed PPTX picture left-to-right across its frame's vertical
    /// axis. PPTX stores this as `a:xfrm/@flipH`; the frame is flipped after
    /// source crop and geometry clipping, before its rotation (issue #1017).
    pub flip_h: bool,
    /// Mirror a fixed PPTX picture top-to-bottom across its frame's horizontal
    /// axis (`a:xfrm/@flipV`; issue #1017).
    pub flip_v: bool,
    pub format: ImageFormat,
    pub width: Option<f64>,
    pub height: Option<f64>,
    pub crop: Option<ImageCrop>,
    /// Optional border stroke around the image.
    pub stroke: Option<BorderSide>,
    /// Horizontal placement inherited from the containing paragraph
    /// (flow documents); None renders at the flow default (left).
    pub alignment: Option<Alignment>,
    /// Clip geometry from the picture's `<a:prstGeom>` (crop to shape).
    pub clip_shape: Option<ImageClipShape>,
    /// Outer shadow effect (`a:effectLst/a:outerShdw` on `p:pic`).
    pub shadow: Option<Shadow>,
    /// Vertical gaps declared by the containing paragraph's `w:spacing`
    /// (flow documents). Word advances a picture paragraph by the picture
    /// plus these, so they have to survive the paragraph being dropped.
    pub paragraph_spacing: Option<ImageParagraphSpacing>,
}

/// The `w:spacing` of the paragraph that held an inline picture, in points.
///
/// Kept apart from [`ImageData::alignment`] because a group of pictures in one
/// paragraph shares a single gap above and below rather than one per picture.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct ImageParagraphSpacing {
    pub before: Option<f64>,
    pub after: Option<f64>,
}

/// Supported **preset** picture clip geometries (PowerPoint "crop to shape").
///
/// A custom `a:custGeom` crop is baked into the image's alpha channel instead
/// and never reaches this enum, so it needs no variant here (issue #872).
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ImageClipShape {
    /// Rounded rectangle with the corner radius as a fraction of the
    /// shorter side (PowerPoint's roundRect `adj`, default 1/6 ≈ 0.1667).
    RoundedRect(f64),
    Ellipse,
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

/// One of the preset DrawingML patterns from `ST_PresetPatternVal`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PatternPreset {
    Percent5,
    Percent10,
    Percent20,
    Percent25,
    Percent30,
    Percent40,
    Percent50,
    Percent60,
    Percent70,
    Percent75,
    Percent80,
    Percent90,
    Horizontal,
    Vertical,
    LightHorizontal,
    LightVertical,
    DarkHorizontal,
    DarkVertical,
    NarrowHorizontal,
    NarrowVertical,
    DashedHorizontal,
    DashedVertical,
    Cross,
    DownwardDiagonal,
    UpwardDiagonal,
    LightDownwardDiagonal,
    LightUpwardDiagonal,
    DarkDownwardDiagonal,
    DarkUpwardDiagonal,
    WideDownwardDiagonal,
    WideUpwardDiagonal,
    DashedDownwardDiagonal,
    DashedUpwardDiagonal,
    DiagonalCross,
    SmallCheck,
    LargeCheck,
    SmallGrid,
    LargeGrid,
    DotGrid,
    SmallConfetti,
    LargeConfetti,
    HorizontalBrick,
    DiagonalBrick,
    SolidDiamond,
    OpenDiamond,
    DottedDiamond,
    Plaid,
    Sphere,
    Weave,
    Divot,
    Shingle,
    Wave,
    Trellis,
    ZigZag,
}

impl PatternPreset {
    /// Every preset defined by DrawingML's `ST_PresetPatternVal`.
    pub const ALL: [Self; 54] = [
        Self::Percent5,
        Self::Percent10,
        Self::Percent20,
        Self::Percent25,
        Self::Percent30,
        Self::Percent40,
        Self::Percent50,
        Self::Percent60,
        Self::Percent70,
        Self::Percent75,
        Self::Percent80,
        Self::Percent90,
        Self::Horizontal,
        Self::Vertical,
        Self::LightHorizontal,
        Self::LightVertical,
        Self::DarkHorizontal,
        Self::DarkVertical,
        Self::NarrowHorizontal,
        Self::NarrowVertical,
        Self::DashedHorizontal,
        Self::DashedVertical,
        Self::Cross,
        Self::DownwardDiagonal,
        Self::UpwardDiagonal,
        Self::LightDownwardDiagonal,
        Self::LightUpwardDiagonal,
        Self::DarkDownwardDiagonal,
        Self::DarkUpwardDiagonal,
        Self::WideDownwardDiagonal,
        Self::WideUpwardDiagonal,
        Self::DashedDownwardDiagonal,
        Self::DashedUpwardDiagonal,
        Self::DiagonalCross,
        Self::SmallCheck,
        Self::LargeCheck,
        Self::SmallGrid,
        Self::LargeGrid,
        Self::DotGrid,
        Self::SmallConfetti,
        Self::LargeConfetti,
        Self::HorizontalBrick,
        Self::DiagonalBrick,
        Self::SolidDiamond,
        Self::OpenDiamond,
        Self::DottedDiamond,
        Self::Plaid,
        Self::Sphere,
        Self::Weave,
        Self::Divot,
        Self::Shingle,
        Self::Wave,
        Self::Trellis,
        Self::ZigZag,
    ];

    /// Parse the serialized value of DrawingML's `ST_PresetPatternVal`.
    pub(crate) fn from_ooxml(value: &str) -> Option<Self> {
        Some(match value {
            "pct5" => Self::Percent5,
            "pct10" => Self::Percent10,
            "pct20" => Self::Percent20,
            "pct25" => Self::Percent25,
            "pct30" => Self::Percent30,
            "pct40" => Self::Percent40,
            "pct50" => Self::Percent50,
            "pct60" => Self::Percent60,
            "pct70" => Self::Percent70,
            "pct75" => Self::Percent75,
            "pct80" => Self::Percent80,
            "pct90" => Self::Percent90,
            "horz" => Self::Horizontal,
            "vert" => Self::Vertical,
            "ltHorz" => Self::LightHorizontal,
            "ltVert" => Self::LightVertical,
            "dkHorz" => Self::DarkHorizontal,
            "dkVert" => Self::DarkVertical,
            "narHorz" => Self::NarrowHorizontal,
            "narVert" => Self::NarrowVertical,
            "dashHorz" => Self::DashedHorizontal,
            "dashVert" => Self::DashedVertical,
            "cross" => Self::Cross,
            "dnDiag" => Self::DownwardDiagonal,
            "upDiag" => Self::UpwardDiagonal,
            "ltDnDiag" => Self::LightDownwardDiagonal,
            "ltUpDiag" => Self::LightUpwardDiagonal,
            "dkDnDiag" => Self::DarkDownwardDiagonal,
            "dkUpDiag" => Self::DarkUpwardDiagonal,
            "wdDnDiag" => Self::WideDownwardDiagonal,
            "wdUpDiag" => Self::WideUpwardDiagonal,
            "dashDnDiag" => Self::DashedDownwardDiagonal,
            "dashUpDiag" => Self::DashedUpwardDiagonal,
            "diagCross" => Self::DiagonalCross,
            "smCheck" => Self::SmallCheck,
            "lgCheck" => Self::LargeCheck,
            "smGrid" => Self::SmallGrid,
            "lgGrid" => Self::LargeGrid,
            "dotGrid" => Self::DotGrid,
            "smConfetti" => Self::SmallConfetti,
            "lgConfetti" => Self::LargeConfetti,
            "horzBrick" => Self::HorizontalBrick,
            "diagBrick" => Self::DiagonalBrick,
            "solidDmnd" => Self::SolidDiamond,
            "openDmnd" => Self::OpenDiamond,
            "dotDmnd" => Self::DottedDiamond,
            "plaid" => Self::Plaid,
            "sphere" => Self::Sphere,
            "weave" => Self::Weave,
            "divot" => Self::Divot,
            "shingle" => Self::Shingle,
            "wave" => Self::Wave,
            "trellis" => Self::Trellis,
            "zigZag" => Self::ZigZag,
            _ => return None,
        })
    }
}

/// A DrawingML preset pattern with foreground and background colors.
#[derive(Debug, Clone)]
pub struct PatternFill {
    pub preset: PatternPreset,
    pub foreground: Color,
    pub background: Color,
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

/// A supported DrawingML top/front-face bevel.
///
/// The current renderer models the default circular bevel under an
/// orthographic, top-directed three-point light rig. `width` is how far the
/// bevel reaches into the face; `height` is how far it rises above it. Rendering
/// is currently limited to rectangles with a solid, gradient, or pattern fill;
/// other shape kinds and unfilled shapes carry this value without drawing a
/// rim.
#[derive(Debug, Clone)]
pub struct TopBevel {
    pub width: f64,
    pub height: f64,
    /// Rotation of the light rig around the viewing axis, in degrees.
    pub light_rig_rotation_deg: f64,
}

/// What a chart axis' or gridline's `<c:spPr>` says about its line.
///
/// The three states are distinct, exactly as they are for the chart area
/// ([`ChartAreaOutline`], issue #637): saying nothing means the automatic
/// line, `<a:ln><a:noFill/></a:ln>` means none at all, and a stated `<a:ln>`
/// means that one. The two enums have the same shape and are candidates for
/// unification.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ChartLine {
    /// The part states no `<a:ln>`; the renderer's automatic stroke applies.
    #[default]
    Automatic,
    /// `<a:ln><a:noFill/></a:ln>` — draw nothing.
    Suppressed,
    /// A stated line. Either half may still be absent: a `<a:ln>` naming only
    /// a width keeps the automatic colour, and vice versa.
    Explicit {
        width_pt: Option<f64>,
        color: Option<Color>,
    },
}

/// Basic geometric shape.
#[derive(Debug, Clone)]
pub struct Shape {
    pub kind: ShapeKind,
    pub fill: Option<Color>,
    /// Gradient fill for the shape (takes precedence over solid fill when present).
    pub gradient_fill: Option<GradientFill>,
    /// DrawingML preset pattern fill (takes precedence over gradient and solid fills).
    pub pattern_fill: Option<PatternFill>,
    pub stroke: Option<BorderSide>,
    /// Rotation angle in degrees (clockwise).
    pub rotation_deg: Option<f64>,
    /// Opacity from 0.0 (fully transparent) to 1.0 (fully opaque).
    pub opacity: Option<f64>,
    /// Outer shadow effect.
    pub shadow: Option<Shadow>,
    /// Top/front-face bevel effect (rendered on filled rectangles only).
    pub top_bevel: Option<TopBevel>,
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
    /// Several subpaths filled as one path under the even-odd rule, so an
    /// inner boundary carves a hole rather than painting solid.
    ///
    /// This is what a DrawingML `a:custGeom` is: its `a:pathLst` may hold
    /// separate `a:path` elements, and one `a:path` may hold several subpaths.
    /// Vertices are normalized to 0.0-1.0 of the bounding box, like
    /// [`ShapeKind::Polygon`] (issue #870).
    Path {
        subpaths: Vec<Subpath>,
    },
}

/// One outline of a custom geometry.
///
/// Vertices are normalized to 0.0-1.0 of the shape's bounding box.
#[derive(Debug, Clone, PartialEq)]
pub struct Subpath {
    pub vertices: Vec<(f64, f64)>,
    /// Whether `a:close` ended the outline, so its stroke joins the last
    /// vertex back to the first. An open polyline must not be closed: the
    /// elbow connectors of the deck on issue #1205 are three-segment
    /// `moveTo lnTo lnTo lnTo` paths, and closing them draws a diagonal the
    /// deck never declared. A fill still treats an open outline as closed,
    /// which is what PowerPoint does.
    pub closed: bool,
}

impl Subpath {
    /// An outline that returns to its own start.
    pub fn closed_outline(vertices: Vec<(f64, f64)>) -> Self {
        Self {
            vertices,
            closed: true,
        }
    }

    /// An outline that stops at its last vertex.
    pub fn open_outline(vertices: Vec<(f64, f64)>) -> Self {
        Self {
            vertices,
            closed: false,
        }
    }

    /// A closed outline needs three vertices to enclose an area; an open one
    /// draws a line from two.
    pub(crate) fn encloses_or_draws(&self) -> bool {
        let needed: usize = if self.closed { 3 } else { 2 };
        self.vertices.len() >= needed
    }
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
