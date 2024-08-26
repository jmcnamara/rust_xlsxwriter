// shape - A module to represent Excel cell shapes.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use crate::drawing::{DrawingObject, DrawingType};
use crate::{Format, ObjectMovement, Url};

#[derive(Clone)]
/// The `Shape` struct represents an worksheet shape object.
///
/// A Shape is TODO
///
pub struct Shape {
    height: f64,
    width: f64,
    pub(crate) x_offset: u32,
    pub(crate) y_offset: u32,
    pub(crate) text: String,
    pub(crate) alt_text: String,
    pub(crate) object_movement: ObjectMovement,
    pub(crate) decorative: bool,
    pub(crate) format: Option<Format>,
    pub(crate) url: Option<Url>,
    pub(crate) _shape_type: ShapeType,
}

impl Shape {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new Shape object to represent an Excel cell shape.
    ///
    /// The text of the Shape is added in the constructor.
    ///
    /// # Examples
    ///
    /// TODO
    ///
    pub fn textbox() -> Shape {
        Shape {
            x_offset: 0,
            y_offset: 0,

            width: 192.0,
            height: 120.0,
            text: String::new(),
            alt_text: String::new(),
            object_movement: ObjectMovement::MoveAndSizeWithCells,
            decorative: false,
            format: None,
            url: None,
            _shape_type: ShapeType::TextBox,
        }
    }

    /// Set the text in the shape.
    ///
    /// TODO
    ///
    /// # Parameters
    ///
    /// - `text`: The text for the shape.
    ///
    /// # Examples
    ///
    ///
    pub fn set_text(mut self, text: impl Into<String>) -> Shape {
        self.text = text.into();
        self
    }

    /// Set the width of the shape in pixels.
    ///
    /// The default width of an Excel shape is 128 pixels.
    ///
    /// # Parameters
    ///
    /// - `width`: The shape width in pixels.
    ///
    /// # Examples
    ///
    ///
    /// TODO
    ///
    pub fn set_width(mut self, width: u32) -> Shape {
        if width == 0 {
            return self;
        }

        self.width = f64::from(width);
        self
    }

    /// Set the height of the shape in pixels.
    ///
    /// The default height of an Excel shape is 74 pixels. See the example above.
    ///
    /// # Parameters
    ///
    /// - `height`: The shape height in pixels.
    ///
    pub fn set_height(mut self, height: u32) -> Shape {
        if height == 0 {
            return self;
        }
        self.height = f64::from(height);
        self
    }

    /// Set the [`Format`] of the shape.
    ///
    /// Set the font or background properties of a shape using a [`Format`]
    /// object. Only the font name, size and background color are supported.
    ///
    /// This API is currently experimental and may go away in the future.
    ///
    /// # Parameters
    ///
    /// - `format`: The [`Format`] property for the shape.
    ///
    pub fn set_format(mut self, format: impl Into<Format>) -> Shape {
        self.format = Some(format.into());
        self
    }

    /// Set the alt text for the shape to help accessibility.
    ///
    /// The alt text is used with screen readers to help people with visual
    /// disabilities.
    ///
    /// See the following Microsoft documentation on [Everything you need to
    /// know to write effective alt
    /// text](https://support.microsoft.com/en-us/office/everything-you-need-to-know-to-write-effective-alt-text-df98f884-ca3d-456c-807b-1a1fa82f5dc2).
    ///
    /// # Parameters
    ///
    /// - `alt_text`: The alt text string to add to the shape.
    ///
    pub fn set_alt_text(mut self, alt_text: impl Into<String>) -> Shape {
        let alt_text = alt_text.into();
        if alt_text.chars().count() > 255 {
            eprintln!("Alternative text is greater than Excel's limit of 255 characters.");
            return self;
        }

        self.alt_text = alt_text;
        self
    }

    /// Set the object movement options for a worksheet shape.
    ///
    /// Set the option to define how an shape will behave in Excel if the cells
    /// under the shape are moved, deleted, or have their size changed. In
    /// Excel the options are:
    ///
    /// 1. Move and size with cells.
    /// 2. Move but don't size with cells.
    /// 3. Don't move or size with cells.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/object_movement.png">
    ///
    /// These values are defined in the [`ObjectMovement`] enum.
    ///
    /// The [`ObjectMovement`] enum also provides an additional option to "Move
    /// and size with cells - after the shape is inserted" to allow shapes to
    /// be hidden in rows or columns. In Excel this equates to option 1 above
    /// but the internal shape position calculations are handled differently.
    ///
    /// # Parameters
    ///
    /// - `option`: An shape/object positioning behavior defined by the
    ///   [`ObjectMovement`] enum.
    pub fn set_object_movement(mut self, option: ObjectMovement) -> Shape {
        self.object_movement = option;
        self
    }
}

// TODO
#[derive(Clone, PartialEq, Eq)]
pub(crate) enum ShapeType {
    // TODO
    TextBox,
}

// Trait for objects that have a component stored in the drawing.xml file.
impl DrawingObject for Shape {
    #[allow(clippy::if_same_then_else)]
    fn x_offset(&self) -> u32 {
        self.x_offset
    }

    #[allow(clippy::if_same_then_else)]
    fn y_offset(&self) -> u32 {
        self.y_offset
    }

    fn width_scaled(&self) -> f64 {
        self.width
    }

    fn height_scaled(&self) -> f64 {
        self.height
    }

    fn object_movement(&self) -> ObjectMovement {
        self.object_movement
    }

    fn name(&self) -> String {
        self.text.clone()
    }

    fn alt_text(&self) -> String {
        self.alt_text.clone()
    }

    fn decorative(&self) -> bool {
        self.decorative
    }

    fn drawing_type(&self) -> DrawingType {
        DrawingType::Shape
    }
}
