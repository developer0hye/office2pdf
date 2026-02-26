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
}

/// RGB color.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Color {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl Color {
    pub fn new(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b }
    }

    pub fn black() -> Self {
        Self { r: 0, g: 0, b: 0 }
    }

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
}
