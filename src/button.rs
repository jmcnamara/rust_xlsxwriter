// button - A module for handling Excel button files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use crate::drawing::{DrawingObject, DrawingType};
use crate::vml::VmlInfo;
use crate::{ObjectMovement, DEFAULT_COL_WIDTH_PIXELS, DEFAULT_ROW_HEIGHT_PIXELS};

#[derive(Clone, Debug)]
/// The `Button` struct is used to create an object to represent an button that
/// can be inserted into a worksheet.
///
/// TODO
///
pub struct Button {
    height: f64,
    width: f64,
    scale_width: f64,
    scale_height: f64,
    pub(crate) x_offset: u32,
    pub(crate) y_offset: u32,
    pub(crate) name: String,
    pub(crate) macro_name: String,
    pub(crate) alt_text: String,
    pub(crate) object_movement: ObjectMovement,
    pub(crate) decorative: bool,
}

impl Default for Button {
    fn default() -> Self {
        Self::new()
    }
}

impl Button {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new Button object TODO.
    ///
    pub fn new() -> Button {
        Button {
            height: f64::from(DEFAULT_ROW_HEIGHT_PIXELS),
            width: f64::from(DEFAULT_COL_WIDTH_PIXELS),
            scale_width: 1.0,
            scale_height: 1.0,
            x_offset: 0,
            y_offset: 0,
            name: String::new(),
            macro_name: String::new(),
            alt_text: String::new(),
            object_movement: ObjectMovement::MoveAndSizeWithCells,
            decorative: false,
        }
    }

    /// Set the button caption. TODO.
    ///
    /// TODO
    ///
    /// # Parameters
    ///
    /// `todo` - TODO.
    ///
    pub fn set_caption(mut self, caption: impl Into<String>) -> Button {
        self.name = caption.into();
        self
    }

    /// Set the button caption. TODO.
    ///
    /// TODO
    ///
    /// # Parameters
    ///
    /// `todo` - TODO.
    ///
    pub fn set_macro(mut self, name: impl Into<String>) -> Button {
        self.macro_name = name.into();
        self
    }

    /// Set the height scale for the button relative to 1.0/100%.
    ///
    /// # Parameters
    ///
    /// * `scale` - The scale ratio.
    ///
    /// TODO example
    ///
    pub fn set_scale_height(mut self, scale: f64) -> Button {
        if scale <= 0.0 {
            return self;
        }

        self.scale_height = scale;
        self
    }

    /// Set the width scale for the button relative to 1.0/100%.
    ///
    /// # Parameters
    ///
    /// * `scale` - The scale ratio.
    ///
    pub fn set_scale_width(mut self, scale: f64) -> Button {
        if scale <= 0.0 {
            return self;
        }

        self.scale_width = scale;
        self
    }

    /// Set the width and height scale to achieve a specific size. TODO
    ///
    /// Calculate and set the horizontal and vertical scales for an button in
    /// order to display it at a fixed width and height in a worksheet. This is
    /// most commonly used to scale an button so that it fits within a cell or a
    /// specific region in a worksheet.
    ///
    /// There are two options, which are controlled by the `keep_aspect_ratio`
    /// parameter. The button can be scaled vertically and horizontally to give
    /// the specified with and height or the aspect ratio of the button can be
    /// maintained so that the button is scaled to the lesser of the horizontal
    /// or vertical sizes. See the example below.
    ///
    /// See also the
    /// [`worksheet.insert_button_fit_to_cell()`](crate::Worksheet::insert_button_fit_to_cell)
    /// method.
    ///
    /// # Parameters
    ///
    /// * `width` - The target width in pixels to scale the button to.
    /// * `height` - The target height in pixels to scale the button to.
    /// * `keep_aspect_ratio` - Boolean value to maintain the aspect ratio of
    ///   the button if `true` or scale independently in the horizontal and
    ///   vertical directions if `false`.
    ///
    /// Note: the `width` and `height` can mainly be considered as pixel sizes.
    /// However, f64 values are allowed for cases where a fractional size is
    /// required
    ///
    ///
    pub fn set_scale_to_size<T>(mut self, width: T, height: T, keep_aspect_ratio: bool) -> Button
    where
        T: Into<f64> + Copy,
    {
        if width.into() == 0.0 || height.into() == 0.0 {
            return self;
        }

        let mut scale_width = width.into() / self.width();
        let mut scale_height = height.into() / self.height();

        if keep_aspect_ratio {
            if scale_width < scale_height {
                scale_height = scale_width;
            } else {
                scale_width = scale_height;
            }
        }

        self = self.set_scale_width(scale_width);
        self = self.set_scale_height(scale_height);

        self
    }

    /// Set the alt text for the button. TODO
    ///
    /// Set the alt text for the button to help accessibility. The alt text is
    /// used with screen readers to help people with visual disabilities.
    ///
    /// See the following Microsoft documentation on [Everything you need to
    /// know to write effective alt
    /// text](https://support.microsoft.com/en-us/office/everything-you-need-to-know-to-write-effective-alt-text-df98f884-ca3d-456c-807b-1a1fa82f5dc2).
    ///
    /// # Parameters
    ///
    /// * `alt_text` - The alt text string to add to the button.
    ///
    /// # Examples
    ///
    pub fn set_alt_text(mut self, alt_text: impl Into<String>) -> Button {
        self.alt_text = alt_text.into();
        self
    }

    /// Set the object movement options for a worksheet button.
    ///
    /// Set the option to define how an button will behave in Excel if the cells
    /// under the button are moved, deleted, or have their size changed. In Excel
    /// the options are:
    ///
    /// 1. Move and size with cells.
    /// 2. Move but don't size with cells.
    /// 3. Don't move or size with cells.
    ///
    /// <img src="https://rustxlsxwriter.github.io/buttons/object_movement.png">
    ///
    /// These values are defined in the [`ObjectMovement`] enum.
    ///
    /// The [`ObjectMovement`] enum also provides an additional option to
    /// "Move and size with cells - after the button is inserted" to allow buttons
    /// to be hidden in rows or columns. In Excel this equates to option 1 above
    /// but the internal button position calculations are handled differently.
    ///
    /// # Parameters
    ///
    /// * `option` - An button/object positioning behavior defined by the
    ///   [`ObjectMovement`] enum.
    ///
    /// # Examples
    ///
    ///  TODO
    ///
    pub fn set_object_movement(mut self, option: ObjectMovement) -> Button {
        self.object_movement = option;
        self
    }

    /// Get the width of the button used for the size calculations in Excel.
    ///
    /// TODO.
    ///
    /// # Examples
    ///
    /// This example shows how to get some of the properties of an Button that
    /// will be used in an Excel worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_button_dimensions.rs
    /// #
    /// # use rust_xlsxwriter::{Button, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    ///     let button = Button::new();
    ///
    ///     assert_eq!(106.0, button.width());
    ///     assert_eq!(106.0, button.height());
    ///     assert_eq!(96.0, button.width_dpi());
    ///     assert_eq!(96.0, button.height_dpi());
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    pub fn width(&self) -> f64 {
        self.width
    }

    /// Get the height of the button used for the size calculations in Excel. See
    /// the example above.
    ///
    pub fn height(&self) -> f64 {
        self.height
    }

    // Buttons are stored in a vmlDrawing file. We create a struct to store the
    // required image information in that format.
    pub(crate) fn vml_info(&self) -> VmlInfo {
        VmlInfo {
            width: self.width,
            height: self.height,
            name: self.name.clone(),
            alt_text: self.alt_text.clone(),
            macro_name: self.macro_name.clone(),
            ..Default::default()
        }
    }

    // -----------------------------------------------------------------------
    // Internal methods.
    // -----------------------------------------------------------------------
}

// Trait for objects that have a component stored in the drawing.xml file.
impl DrawingObject for Button {
    fn x_offset(&self) -> u32 {
        self.x_offset
    }

    fn y_offset(&self) -> u32 {
        self.y_offset
    }

    fn width_scaled(&self) -> f64 {
        self.width * self.scale_width
    }

    fn height_scaled(&self) -> f64 {
        self.height * self.scale_height
    }

    fn object_movement(&self) -> ObjectMovement {
        self.object_movement
    }

    fn name(&self) -> String {
        self.name.clone()
    }

    fn alt_text(&self) -> String {
        self.alt_text.clone()
    }

    fn decorative(&self) -> bool {
        self.decorative
    }

    fn drawing_type(&self) -> DrawingType {
        DrawingType::Button
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs/functions.
// -----------------------------------------------------------------------
