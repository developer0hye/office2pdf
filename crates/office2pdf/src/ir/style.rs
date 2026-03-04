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
#[derive(Debug, Clone, Default)]
pub struct TextStyle {
    pub font_family: Option<String>,
    pub font_size: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strikethrough: Option<bool>,
    pub color: Option<Color>,
    /// Superscript or subscript vertical alignment.
    pub vertical_align: Option<VerticalTextAlign>,
    /// All caps: render text in uppercase.
    pub all_caps: Option<bool>,
    /// Small caps: render lowercase letters as smaller uppercase.
    pub small_caps: Option<bool>,
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
mod tests {
    use super::*;

    #[test]
    fn test_color_constructors() {
        let black = Color::black();
        assert_eq!(black, Color { r: 0, g: 0, b: 0 });
        let white = Color::white();
        assert_eq!(
            white,
            Color {
                r: 255,
                g: 255,
                b: 255,
            }
        );
    }

    #[test]
    fn test_stylesheet_default_is_empty() {
        let ss = StyleSheet::default();
        assert!(ss.styles.is_empty());
    }

    #[test]
    fn test_text_style_default_has_none_text_effects() {
        let ts = TextStyle::default();
        assert!(ts.vertical_align.is_none());
        assert!(ts.all_caps.is_none());
        assert!(ts.small_caps.is_none());
    }

    #[test]
    fn test_text_style_superscript() {
        let ts = TextStyle {
            vertical_align: Some(VerticalTextAlign::Superscript),
            ..TextStyle::default()
        };
        assert_eq!(ts.vertical_align, Some(VerticalTextAlign::Superscript));
    }

    #[test]
    fn test_text_style_subscript() {
        let ts = TextStyle {
            vertical_align: Some(VerticalTextAlign::Subscript),
            ..TextStyle::default()
        };
        assert_eq!(ts.vertical_align, Some(VerticalTextAlign::Subscript));
    }

    #[test]
    fn test_text_style_caps() {
        let ts = TextStyle {
            all_caps: Some(true),
            small_caps: Some(true),
            ..TextStyle::default()
        };
        assert_eq!(ts.all_caps, Some(true));
        assert_eq!(ts.small_caps, Some(true));
    }
}
