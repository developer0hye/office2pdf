mod document;
mod elements;
mod style;

pub use document::*;
pub use elements::*;
pub use style::*;

/// Common interface for elements with position and dimensions.
pub trait Positioned {
    /// X position in points.
    fn x(&self) -> f64;
    /// Y position in points.
    fn y(&self) -> f64;
    /// Width in points.
    fn width(&self) -> f64;
    /// Height in points.
    fn height(&self) -> f64;
}

impl Positioned for FixedElement {
    fn x(&self) -> f64 {
        self.x
    }

    fn y(&self) -> f64 {
        self.y
    }

    fn width(&self) -> f64 {
        self.width
    }

    fn height(&self) -> f64 {
        self.height
    }
}

impl Positioned for FloatingImage {
    fn x(&self) -> f64 {
        self.offset_x
    }

    fn y(&self) -> f64 {
        self.offset_y
    }

    fn width(&self) -> f64 {
        self.image.width.unwrap_or(0.0)
    }

    fn height(&self) -> f64 {
        self.image.height.unwrap_or(0.0)
    }
}

impl Positioned for FloatingTextBox {
    fn x(&self) -> f64 {
        self.offset_x
    }

    fn y(&self) -> f64 {
        self.offset_y
    }

    fn width(&self) -> f64 {
        self.width
    }

    fn height(&self) -> f64 {
        self.height
    }
}

#[cfg(test)]
#[path = "positioned_tests.rs"]
mod positioned_tests;
