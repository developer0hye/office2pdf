/// Collection of named styles in the document.
#[derive(Debug, Clone, Default)]
pub struct StyleSheet {
    pub styles: Vec<NamedStyle>,
}

/// A named style that can be referenced by paragraphs/runs.
#[derive(Debug, Clone)]
pub struct NamedStyle {
    pub id: String,
    pub name: String,
    pub paragraph: Option<ParagraphStyle>,
    pub text: Option<TextStyle>,
}

/// Paragraph-level formatting.
#[derive(Debug, Clone, Default)]
pub struct ParagraphStyle {
    pub alignment: Option<Alignment>,
    pub indent_left: Option<f64>,
    pub indent_right: Option<f64>,
    pub indent_first_line: Option<f64>,
    pub line_spacing: Option<LineSpacing>,
    pub space_before: Option<f64>,
    pub space_after: Option<f64>,
    /// Heading level (1 = H1, 2 = H2, ..., 6 = H6). When set, the paragraph
    /// is emitted as a Typst `#heading` element for proper PDF structure tagging.
    pub heading_level: Option<u8>,
    /// Text direction for bidirectional rendering (RTL for Arabic/Hebrew).
    pub direction: Option<TextDirection>,
    /// Custom tab stop positions for this paragraph.
    pub tab_stops: Option<Vec<TabStop>>,
}

/// A custom tab stop definition.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TabStop {
    /// Position in points from the left margin.
    pub position: f64,
    /// Alignment of text at this tab stop.
    pub alignment: TabAlignment,
    /// Leader character filling the space before this tab stop.
    pub leader: TabLeader,
}

/// Tab stop alignment type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TabAlignment {
    #[default]
    Left,
    Center,
    Right,
    Decimal,
}

/// Leader character for a tab stop.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum TabLeader {
    #[default]
    None,
    Dot,
    Hyphen,
    Underscore,
}

/// Text direction for bidirectional (BiDi) rendering.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextDirection {
    /// Left-to-right (default for Latin, CJK scripts).
    Ltr,
    /// Right-to-left (Arabic, Hebrew scripts).
    Rtl,
}

/// Text alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
}

/// Line spacing specification.
#[derive(Debug, Clone, Copy)]
pub enum LineSpacing {
    /// Multiplier (e.g. 1.0 = single, 1.5, 2.0 = double).
    Proportional(f64),
    /// Exact spacing in points.
    Exact(f64),
}

/// Vertical alignment for superscript/subscript text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalTextAlign {
    Superscript,
    Subscript,
}

/// Character-level formatting.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct TextStyle {
    pub font_family: Option<String>,
    pub font_size: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strikethrough: Option<bool>,
    pub color: Option<Color>,
    /// Text highlight background color.
    pub highlight: Option<Color>,
    /// Superscript or subscript vertical alignment.
    pub vertical_align: Option<VerticalTextAlign>,
    /// All caps: render text in uppercase.
    pub all_caps: Option<bool>,
    /// Small caps: render lowercase letters as smaller uppercase.
    pub small_caps: Option<bool>,
    /// Character spacing (letter spacing / tracking) in points.
    pub letter_spacing: Option<f64>,
}

impl TextStyle {
    /// Merge fields from `other` into `self`. For each field, if `other` has
    /// `Some(value)`, it overwrites `self`'s value. Fields that are `None` in
    /// `other` are left unchanged.
    pub fn merge_from(&mut self, other: &TextStyle) {
        if other.font_family.is_some() {
            self.font_family = other.font_family.clone();
        }
        if other.font_size.is_some() {
            self.font_size = other.font_size;
        }
        if other.bold.is_some() {
            self.bold = other.bold;
        }
        if other.italic.is_some() {
            self.italic = other.italic;
        }
        if other.underline.is_some() {
            self.underline = other.underline;
        }
        if other.strikethrough.is_some() {
            self.strikethrough = other.strikethrough;
        }
        if other.color.is_some() {
            self.color = other.color;
        }
        if other.highlight.is_some() {
            self.highlight = other.highlight;
        }
        if other.vertical_align.is_some() {
            self.vertical_align = other.vertical_align;
        }
        if other.all_caps.is_some() {
            self.all_caps = other.all_caps;
        }
        if other.small_caps.is_some() {
            self.small_caps = other.small_caps;
        }
        if other.letter_spacing.is_some() {
            self.letter_spacing = other.letter_spacing;
        }
    }
}

impl ParagraphStyle {
    /// Merge fields from `other` into `self`. For each field, if `other` has
    /// `Some(value)`, it overwrites `self`'s value. Fields that are `None` in
    /// `other` are left unchanged.
    pub fn merge_from(&mut self, other: &ParagraphStyle) {
        if other.alignment.is_some() {
            self.alignment = other.alignment;
        }
        if other.indent_left.is_some() {
            self.indent_left = other.indent_left;
        }
        if other.indent_right.is_some() {
            self.indent_right = other.indent_right;
        }
        if other.indent_first_line.is_some() {
            self.indent_first_line = other.indent_first_line;
        }
        if other.line_spacing.is_some() {
            self.line_spacing = other.line_spacing;
        }
        if other.space_before.is_some() {
            self.space_before = other.space_before;
        }
        if other.space_after.is_some() {
            self.space_after = other.space_after;
        }
        if other.heading_level.is_some() {
            self.heading_level = other.heading_level;
        }
        if other.direction.is_some() {
            self.direction = other.direction;
        }
        if other.tab_stops.is_some() {
            self.tab_stops = other.tab_stops.clone();
        }
    }
}

/// RGB color.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Color {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl Color {
    /// Create a color from RGB components.
    pub fn new(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b }
    }

    /// Black (`#000000`).
    pub fn black() -> Self {
        Self { r: 0, g: 0, b: 0 }
    }

    /// White (`#FFFFFF`).
    pub fn white() -> Self {
        Self {
            r: 255,
            g: 255,
            b: 255,
        }
    }
}

#[cfg(test)]
#[path = "style_tests.rs"]
mod tests;
