// button - A module for handling Excel button files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use crate::drawing::{DrawingObject, DrawingType};
use crate::vml::VmlInfo;
use crate::{ObjectMovement, DEFAULT_COL_WIDTH_PIXELS, DEFAULT_ROW_HEIGHT_PIXELS};

#[derive(Clone)]
/// The `Button` struct represents an worksheet button object.
///
/// The `Button` struct is used to create an Excel "Form Control" button object
/// to represent a button on a worksheet.
///
/// <img src="https://rustxlsxwriter.github.io/images/doc_button_intro.png">
///
/// The worksheet button object is mainly provided as a way of triggering a VBA
/// macro, see [Working with VBA macros](crate::macros) for more details. It is
/// used in conjunction with the
/// [`Worksheet::insert_button()`](crate::Worksheet::insert_button) method.
///
/// Note, Button is the only VBA Control supported by `rust_xlsxwriter`. It is
/// unlikely that any other Excel form elements will be added in the future due
/// to the implementation effort required.
///
/// Here is a complete example with a button that has a macro attached to it.
///
/// ```
/// # // This code is available in examples/app_macros.rs
/// #
/// use rust_xlsxwriter::{Button, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add the VBA macro file.
///     workbook.add_vba_project("examples/vbaProject.bin")?;
///
///     // Add a worksheet and some text.
///     let worksheet = workbook.add_worksheet();
///
///     // Widen the first column for clarity.
///     worksheet.set_column_width(0, 30)?;
///
///     worksheet.write(2, 0, "Press the button to say hello:")?;
///
///     // Add a button tied to a macro in the VBA project.
///     let button = Button::new()
///         .set_caption("Press Me")
///         .set_macro("say_hello")
///         .set_width(80)
///         .set_height(30);
///
///     worksheet.insert_button(2, 1, &button)?;
///
///     // Save the file to disk. Note the `.xlsm` extension. This is required by
///     // Excel or it raise a warning.
///     workbook.save("macros.xlsm")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/app_macros.png">
///
pub struct Button {
    height: f64,
    width: f64,
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

    /// Create a new Button object to represent an Excel Form Control button.
    ///
    pub fn new() -> Button {
        Button {
            x_offset: 0,
            y_offset: 0,
            width: f64::from(DEFAULT_COL_WIDTH_PIXELS),
            height: f64::from(DEFAULT_ROW_HEIGHT_PIXELS),
            name: String::new(),
            alt_text: String::new(),
            macro_name: String::new(),
            object_movement: ObjectMovement::MoveAndSizeWithCells,
            decorative: false,
        }
    }

    /// Set the button caption.
    ///
    /// The default button caption in Excel is "Button 1", "Button 2" etc. This
    /// method can be used to change that caption to some other text.
    ///
    /// # Parameters
    ///
    /// `caption` - The text to display on the button. It must be less than or
    /// equal to 255 characters.
    ///
    /// # Examples
    ///
    /// An example of adding an Excel Form Control button to a worksheet. This
    /// example demonstrates setting the button caption.
    ///
    /// ```
    /// # // This code is available in examples/doc_button_set_caption.rs
    /// #
    /// # use rust_xlsxwriter::{Button, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Add a button with a default caption.
    ///     let button = Button::new();
    ///     worksheet.insert_button(2, 1, &button)?;
    ///
    ///     // Add a button with a user defined caption.
    ///     let button = Button::new().set_caption("Press Me");
    ///     worksheet.insert_button(4, 1, &button)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("button.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/button_set_caption.png">
    ///
    pub fn set_caption(mut self, caption: impl Into<String>) -> Button {
        let caption = caption.into();
        if caption.chars().count() > 255 {
            eprintln!("Button caption is greater than Excel's limit of 255 characters.");
            return self;
        }

        self.name = caption;
        self
    }

    /// Set the macro associated with the button.
    ///
    /// The `set_macro()` method can be used to associate an existing VBA macro
    /// with a button object. See [Working with VBA macros](crate::macros) for
    /// more details on macros in `rust_xlsxwriter`.
    ///
    /// # Parameters
    ///
    /// `name` - The macro name. It should be the same as it appears in the
    /// Excel macros dialog.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/button_macro_dialog.png">
    ///
    ///
    /// # Examples
    ///
    /// An example of adding an Excel Form Control button to a worksheet. This
    /// example demonstrates setting the button macro.
    ///
    /// ```
    /// # // This code is available in examples/doc_button_set_macro.rs
    /// #
    /// # use rust_xlsxwriter::{Button, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add the VBA macro file.
    /// #     workbook.add_vba_project("examples/vbaProject.bin")?;
    /// #
    /// #     // Add a worksheet and some text.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Add a button tied to a macro in the VBA project.
    ///     let button = Button::new().set_macro("say_hello");
    ///
    ///     worksheet.insert_button(2, 1, &button)?;
    /// #
    /// #     // Save the file to disk. Note the `.xlsm` extension.
    /// #     workbook.save("macros.xlsm")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    pub fn set_macro(mut self, name: impl Into<String>) -> Button {
        self.macro_name = name.into();
        self
    }

    /// Set the width of the button in pixels.
    ///
    /// # Parameters
    ///
    /// - `width`: The button width in pixels.
    ///
    pub fn set_width(mut self, width: u32) -> Button {
        if width == 0 {
            return self;
        }

        self.width = f64::from(width);
        self
    }

    /// Set the height of the button in pixels.
    ///
    /// # Parameters
    ///
    /// - `height`: The button height in pixels.
    ///
    pub fn set_height(mut self, height: u32) -> Button {
        if height == 0 {
            return self;
        }
        self.height = f64::from(height);
        self
    }

    /// Set the alt text for the button to help accessibility.
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
    /// - `alt_text`: The alt text string to add to the button.
    ///
    pub fn set_alt_text(mut self, alt_text: impl Into<String>) -> Button {
        let alt_text = alt_text.into();
        if alt_text.chars().count() > 255 {
            eprintln!("Alternative text is greater than Excel's limit of 255 characters.");
            return self;
        }

        self.alt_text = alt_text;
        self
    }

    /// Set the object movement options for a worksheet button.
    ///
    /// Set the option to define how an button will behave in Excel if the cells
    /// under the button are moved, deleted, or have their size changed. In
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
    /// and size with cells - after the button is inserted" to allow buttons to
    /// be hidden in rows or columns. In Excel this equates to option 1 above
    /// but the internal button position calculations are handled differently.
    ///
    /// # Parameters
    ///
    /// - `option`: An button/object positioning behavior defined by the
    ///   [`ObjectMovement`] enum.
    pub fn set_object_movement(mut self, option: ObjectMovement) -> Button {
        self.object_movement = option;
        self
    }

    // Buttons are stored in a vmlDrawing file. We create a struct to store the
    // required image information in that format.
    pub(crate) fn vml_info(&self) -> VmlInfo {
        VmlInfo {
            width: self.width,
            height: self.height,
            text: self.name.clone(),
            alt_text: self.alt_text.clone(),
            macro_name: self.macro_name.clone(),
            fill_color: "buttonFace [67]".to_string(),
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
        self.width
    }

    fn height_scaled(&self) -> f64 {
        self.height
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
        DrawingType::Vml
    }
}
