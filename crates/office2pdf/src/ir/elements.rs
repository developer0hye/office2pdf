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
    List(List),
    PageBreak,
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

/// A table cell.
#[derive(Debug, Clone)]
pub struct TableCell {
    pub content: Vec<Block>,
    pub col_span: u32,
    pub row_span: u32,
    pub border: Option<CellBorder>,
    pub background: Option<Color>,
}

impl Default for TableCell {
    fn default() -> Self {
        Self {
            content: Vec::new(),
            col_span: 1,
            row_span: 1,
            border: None,
            background: None,
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

/// Basic geometric shape.
#[derive(Debug, Clone)]
pub struct Shape {
    pub kind: ShapeKind,
    pub fill: Option<Color>,
    pub stroke: Option<BorderSide>,
    /// Rotation angle in degrees (clockwise).
    pub rotation_deg: Option<f64>,
    /// Opacity from 0.0 (fully transparent) to 1.0 (fully opaque).
    pub opacity: Option<f64>,
}

/// Shape types.
#[derive(Debug, Clone)]
pub enum ShapeKind {
    Rectangle,
    Ellipse,
    Line { x2: f64, y2: f64 },
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
