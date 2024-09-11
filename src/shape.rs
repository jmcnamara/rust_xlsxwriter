// shape - A module to represent Excel cell shapes.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::fmt;

use crate::drawing::{DrawingObject, DrawingType};
use crate::{Color, Formula, ObjectMovement, Url, XlsxError};

#[derive(Clone)]
/// The `Shape` struct represents an worksheet shape object.
///
/// Currently the only Excel shape type that is implemented is the `Textbox`
/// shape:
///
/// ```
/// # // This code is available in examples/app_textbox.rs
/// #
/// use rust_xlsxwriter::{Shape, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Some text to add to the text box.
///     let text = "This is an example of adding a textbox with some text in it";
///
///     // Create a textbox shape and add the text.
///     let textbox = Shape::textbox().set_text(text);
///
///     // Insert a textbox in a cell.
///     worksheet.insert_shape(1, 1, &textbox)?;
///
///     // Save the file to disk.
///     workbook.save("textbox.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/app_textbox.png">
///
/// See also the [`Worksheet::insert_shape()`](crate::Worksheet::insert_shape)
/// and
/// [`Worksheet::insert_shape_with_offset()`](crate::Worksheet::insert_shape_with_offset)
/// methods. Note, it isn't possible to insert textboxes into other
/// `rust_xlsxwriter` objects such as [`Chart`](crate::Chart).
///
/// ## Support for other Excel shape types
///
/// Currently the only Excel shape type that is supported is the `Textbox`
/// shape.
///
/// The internal structure of [`Shape`] and the associated XML generating code
/// is structured to support other shape types but none are currently
/// implemented. The rationale for this is:
///
/// - Unlike applications like `PowerPoint` the shape object is not widely used
///   in Excel.
/// - The most common shape used in Excel is the Textbox/Rectangle.
/// - Alternative ways of displaying information such as [`Image`](crate::Image)
///   or [`Note`](crate::Note) are already supported.
/// - Each shape or connector type requires a significant number of test cases
///   to verify their functionality and interaction.
///
/// The last is the main reason that I don't wish to support other shape types.
/// The implementation burden is small but the test and maintenance burden is
/// large. As such I won't accept Pull Requests to add more shape types.
/// However, I will leave the door open for feature requests that provide a
/// justification.
///
pub struct Shape {
    height: f64,
    width: f64,
    pub(crate) x_offset: u32,
    pub(crate) y_offset: u32,
    pub(crate) text: String,
    pub(crate) text_link: Option<Formula>,
    pub(crate) alt_text: String,
    pub(crate) object_movement: ObjectMovement,
    pub(crate) decorative: bool,
    pub(crate) format: ShapeFormat,
    pub(crate) font: ShapeFont,
    pub(crate) text_options: ShapeText,
    pub(crate) url: Option<Url>,
    pub(crate) _shape_type: ShapeType,
}

impl Shape {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new Shape object to represent an Excel Textbox shape.
    ///
    /// See also the
    /// [`Worksheet::insert_shape()`](crate::Worksheet::insert_shape) and
    /// [`Worksheet::insert_shape_with_offset()`](crate::Worksheet::insert_shape_with_offset)
    /// methods. Note, it isn't possible to insert textboxes into other
    /// `rust_xlsxwriter` objects such as [`Chart`](crate::Chart).
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// text option properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_text_options_set_direction.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeText, ShapeTextDirection, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape and add some text.
    ///     let textbox = Shape::textbox()
    ///         .set_text("古池や\n蛙飛び込む\n水の音")
    ///         .set_text_options(&ShapeText::new()
    ///             .set_direction(ShapeTextDirection::Rotate90EastAsian));
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_options_set_direction.png">
    ///
    pub fn textbox() -> Shape {
        Shape {
            x_offset: 0,
            y_offset: 0,

            width: 192.0,
            height: 120.0,
            text: String::new(),
            text_link: None,
            alt_text: String::new(),
            object_movement: ObjectMovement::MoveAndSizeWithCells,
            decorative: false,
            format: ShapeFormat::default(),
            font: ShapeFont::default(),
            text_options: ShapeText::default(),
            url: None,
            _shape_type: ShapeType::TextBox,
        }
    }

    /// Set the text in the shape.
    ///
    /// This only applies to shapes that have a textbox option.
    ///
    /// See also [`Shape::set_font()`] and [`Shape::set_text_options()`] for
    /// formatting options for text.
    ///
    /// # Parameters
    ///
    /// - `text`: The text for the shape.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape with text to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_set_text.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape and add some text.
    ///     let textbox = Shape::textbox().set_text("This is some text");
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_set_text.png">
    ///
    pub fn set_text(mut self, text: impl Into<String>) -> Shape {
        self.text = text.into();
        self
    }

    /// Set the text in the shape from a worksheet cell.
    ///
    /// Set the textbox text from a link to a worksheet cell like `=A1` or
    /// `=Sheet2!A1`.
    ///
    /// This only applies to shapes that have a textbox option.
    ///
    /// # Parameters
    ///
    /// - `cell`: The cell from which the text is linked. Should be a simple
    ///   string or a [`Formula`].
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape with text from a cell
    /// to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_set_text_link.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Write some text to a cell.
    ///     worksheet.write(1, 5, "This is some text")?;
    ///
    ///     // Create a textbox shape and add some text from a cell.
    ///     let textbox = Shape::textbox().set_text_link("=F2");
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/shape_set_text_link.png">
    ///
    pub fn set_text_link(mut self, cell: impl Into<Formula>) -> Shape {
        self.text_link = Some(cell.into());
        self
    }

    /// Set the width of the shape in pixels.
    ///
    /// The default width for an Excel shape created by `rust_xlsxwriter` is 192
    /// pixels.
    ///
    /// # Parameters
    ///
    /// - `width`: The shape width in pixels. Values less than 5 pixels are
    ///   ignored.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a resized Textbox shape to a worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_set_width.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #   // Create a textbox shape.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_width(100)
    ///         .set_height(100);
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_set_width.png">
    ///
    pub fn set_width(mut self, width: u32) -> Shape {
        if width < 5 {
            return self;
        }

        self.width = f64::from(width);
        self
    }

    /// Set the height of the shape in pixels.
    ///
    /// The default height for an Excel shape created by `rust_xlsxwriter` is
    /// 120 pixels.
    ///
    /// See the example above.
    ///
    /// # Parameters
    ///
    /// - `height`: The shape height in pixels. Values less than 5 pixels are
    ///   ignored.
    ///
    pub fn set_height(mut self, height: u32) -> Shape {
        if height < 5 {
            return self;
        }
        self.height = f64::from(height);
        self
    }

    /// Set the formatting properties for a shape.
    ///
    /// Set the formatting properties for a shape via a [`ShapeFormat`]
    /// object or a sub-struct that implements [`IntoShapeFormat`].
    ///
    /// The formatting that can be applied via a [`ShapeFormat`] object are:
    ///
    /// - [`ShapeFormat::set_solid_fill()`]: Set the [`ShapeSolidFill`] properties.
    /// - [`ShapeFormat::set_pattern_fill()`]: Set the [`ShapePatternFill`] properties.
    /// - [`ShapeFormat::set_gradient_fill()`]: Set the [`ShapeGradientFill`] properties.
    /// - [`ShapeFormat::set_no_fill()`]: Turn off the fill for the shape object.
    /// - [`ShapeFormat::set_line()`]: Set the [`ShapeLine`] properties.
    ///   A synonym for [`ShapeLine`] depending on context.
    /// - [`ShapeFormat::set_no_line()`]: Turn off the line for the shape object.
    ///
    /// # Parameters
    ///
    /// `format`: A [`ShapeFormat`] struct reference or a sub-struct that will
    /// convert into a `ShapeFormat` instance. See the docs for
    /// [`IntoShapeFormat`] for details.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of its
    /// properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_format.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeFormat, ShapeLine, ShapeLineDashType, ShapeSolidFill, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox().set_text("This is some text").set_format(
    ///         &ShapeFormat::new()
    ///             .set_solid_fill(
    ///                 &ShapeSolidFill::new()
    ///                     .set_color("#8ED154")
    ///                     .set_transparency(50),
    ///             )
    ///             .set_line(
    ///                 &ShapeLine::new()
    ///                     .set_color("#FF0000")
    ///                     .set_dash_type(ShapeLineDashType::DashDot),
    ///             ),
    ///     );
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_format.png">
    ///
    pub fn set_format<T>(mut self, format: T) -> Shape
    where
        T: IntoShapeFormat,
    {
        self.format = format.new_shape_format();
        self
    }

    /// Set the font properties of of the shape.
    ///
    /// Set the font properties of a shape text using a [`ShapeFont`] reference.
    /// Example font properties that can be set are:
    ///
    /// - [`ShapeFont::set_bold()`]
    /// - [`ShapeFont::set_italic()`]
    /// - [`ShapeFont::set_color()`]
    /// - [`ShapeFont::set_name()`]
    /// - [`ShapeFont::set_size()`]
    /// - [`ShapeFont::set_underline()`]
    /// - [`ShapeFont::set_strikethrough()`]
    /// - [`ShapeFont::set_right_to_left()`]
    ///
    /// See [`ShapeFont`] for full details.
    ///
    /// # Parameters
    ///
    /// `font`: A [`ShapeFont`] struct reference to represent the font
    /// properties.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// font properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_set_font.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFont, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox().set_text("This is some text").set_font(
    ///         &ShapeFont::new()
    ///             .set_bold()
    ///             .set_italic()
    ///             .set_name("American Typewriter")
    ///             .set_color("#0000FF")
    ///             .set_size(15),
    ///     );
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_set_font.png">
    ///
    pub fn set_font(mut self, font: &ShapeFont) -> Shape {
        self.font = font.clone();
        self
    }

    /// Set the text option properties of of the shape.
    ///
    /// # Parameters
    ///
    /// - `text_options`: The [`ShapeText`] options.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// text option properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_text_options_set_direction.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeText, ShapeTextDirection, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape and add some text.
    ///     let textbox = Shape::textbox()
    ///         .set_text("古池や\n蛙飛び込む\n水の音")
    ///         .set_text_options(&ShapeText::new()
    ///             .set_direction(ShapeTextDirection::Rotate90EastAsian));
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_options_set_direction.png">
    ///
    pub fn set_text_options(mut self, text_options: &ShapeText) -> Shape {
        self.text_options = text_options.clone();
        self
    }

    /// Set a Url/Hyperlink for a shape.
    ///
    /// Set a Url/Hyperlink for an shape so that when the user clicks on it they
    /// are redirected to an internal or external location.
    ///
    /// See [`Url`] for an explanation of the URIs supported by Excel and for
    /// other options that can be set.
    ///
    /// # Parameters
    ///
    /// - `link`: The url/hyperlink associate with the shape as a string or
    ///   [`Url`].
    ///
    /// # Errors
    ///
    /// - [`XlsxError::MaxUrlLengthExceeded`] - URL string or anchor exceeds
    ///   Excel's limit of 2080 characters.
    /// - [`XlsxError::UnknownUrlType`] - The URL has an unknown URI type. See
    ///   [`Worksheet::write_url()`](crate::Worksheet::write_url).
    /// - [`XlsxError::ParameterError`] - URL mouseover tool tip exceeds Excel's
    ///   limit of 255 characters.
    ///
    pub fn set_url(mut self, link: impl Into<Url>) -> Result<Shape, XlsxError> {
        let mut url = link.into();
        url.initialize()?;

        self.url = Some(url);

        Ok(self)
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

// -----------------------------------------------------------------------
// ShapeFormat
// -----------------------------------------------------------------------

#[derive(Clone, PartialEq)]
/// The `ShapeFormat` struct represents formatting for various shape objects.
///
/// Excel uses a standard formatting dialog for the shape elements which
/// generally looks like this:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_format_dialog.png">
///
/// The [`ShapeFormat`] struct represents many of these format options and just
/// like Excel it offers a similar formatting interface for a number of the
/// shape sub-elements supported by `rust_xlsxwriter`.
///
/// It is used in conjunction with the [`Shape`] struct.
///
/// The [`ShapeFormat`] struct is generally passed to the `set_format()` method
/// of a shape element. It supports several shape formatting elements such as:
///
/// - [`ShapeFormat::set_solid_fill()`]: Set the [`ShapeSolidFill`] properties.
/// - [`ShapeFormat::set_pattern_fill()`]: Set the [`ShapePatternFill`]
///   properties.
/// - [`ShapeFormat::set_gradient_fill()`]: Set the [`ShapeGradientFill`]
///   properties.
/// - [`ShapeFormat::set_no_fill()`]: Turn off the fill for the shape object.
/// - [`ShapeFormat::set_line()`]: Set the [`ShapeLine`] properties.
/// - [`ShapeFormat::set_no_line()`]: Turn off the line for the shape object.
///
///
/// # Examples
///
/// This example demonstrates adding a Textbox shape and setting some of its
/// properties.
///
/// ```
/// # // This code is available in examples/doc_shape_format.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Shape, ShapeFormat, ShapeLine, ShapeLineDashType, ShapeSolidFill, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet to the workbook.
/// #     let worksheet = workbook.add_worksheet();
/// #
///     // Create a textbox shape with formatting.
///     let textbox = Shape::textbox().set_text("This is some text").set_format(
///         &ShapeFormat::new()
///             .set_solid_fill(
///                 &ShapeSolidFill::new()
///                     .set_color("#8ED154")
///                     .set_transparency(50),
///             )
///             .set_line(
///                 &ShapeLine::new()
///                     .set_color("#FF0000")
///                     .set_dash_type(ShapeLineDashType::DashDot),
///             ),
///     );
///
///     // Insert a textbox in a cell.
///     worksheet.insert_shape(1, 1, &textbox)?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("shape.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_format.png">
///
pub struct ShapeFormat {
    pub(crate) no_fill: bool,
    pub(crate) no_line: bool,
    pub(crate) line: Option<ShapeLine>,
    pub(crate) solid_fill: Option<ShapeSolidFill>,
    pub(crate) pattern_fill: Option<ShapePatternFill>,
    pub(crate) gradient_fill: Option<ShapeGradientFill>,
}

impl Default for ShapeFormat {
    fn default() -> Self {
        Self::new()
    }
}

impl ShapeFormat {
    /// Create a new `ShapeFormat` instance to set formatting for a shape element.
    ///
    pub fn new() -> ShapeFormat {
        ShapeFormat {
            no_fill: false,
            no_line: false,
            line: None,
            solid_fill: None,
            pattern_fill: None,
            gradient_fill: None,
        }
    }

    /// Set the line formatting for a shape element.
    ///
    /// See the [`ShapeLine`] struct for details on the line properties that can
    /// be set.
    ///
    /// # Parameters
    ///
    /// - `line`: A [`ShapeLine`] struct reference.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// line properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_format_set_line.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeLine, ShapeLineDashType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox().set_text("This is some text").set_format(
    ///         &ShapeFormat::new().set_line(
    ///             &ShapeLine::new()
    ///                 .set_color("#FF0000")
    ///                 .set_dash_type(ShapeLineDashType::DashDot),
    ///         ),
    ///     );
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line.png">
    ///
    pub fn set_line(mut self, line: &ShapeLine) -> ShapeFormat {
        self.line = Some(line.clone());
        self
    }

    /// Turn off the line property for a shape element.
    ///
    /// The line property for a shape element can be turned off if you wish to
    /// hide it.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and turning off its
    /// border.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_format_set_no_line.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFormat, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_format(&ShapeFormat::new().set_no_line());
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_set_hidden.png">
    ///
    pub fn set_no_line(mut self) -> ShapeFormat {
        self.no_line = true;
        self
    }

    /// Set the solid fill formatting for a shape element.
    ///
    /// See the [`ShapeSolidFill`] struct for details on the solid fill
    /// properties that can be set.
    ///
    /// # Parameters
    ///
    /// - `fill`: A [`ShapeSolidFill`] struct reference.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// solid fill properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_solid_fill_set_transparency.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeSolidFill, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox().set_text("This is some text").set_format(
    ///         &ShapeSolidFill::new()
    ///             .set_color("#8ED154")
    ///             .set_transparency(50),
    ///     );
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_solid_fill_set_transparency.png">
    ///
    pub fn set_solid_fill(mut self, fill: &ShapeSolidFill) -> ShapeFormat {
        self.solid_fill = Some(fill.clone());
        self
    }

    /// Turn off the fill property for a shape element.
    ///
    /// The fill property for a shape element can be turned off if you wish to
    /// hide it and display only the border line.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and turning off its
    /// border.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_format_set_no_fill.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFormat, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_format(&ShapeFormat::new().set_no_fill());
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_format_set_no_fill.png">
    ///
    pub fn set_no_fill(mut self) -> ShapeFormat {
        self.no_fill = true;
        self
    }

    /// Set the pattern fill formatting for a shape element.
    ///
    /// See the [`ShapePatternFill`] struct for details on the pattern fill
    /// properties that can be set.
    ///
    /// # Parameters
    ///
    /// - `fill`: A [`ShapePatternFill`] struct reference.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// pattern fill properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_pattern_fill.rs
    /// #
    /// # use rust_xlsxwriter::{Color, Shape, ShapePatternFill, ShapePatternFillType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox().set_text("This is some text").set_format(
    ///         &ShapePatternFill::new()
    ///             .set_pattern(ShapePatternFillType::Dotted20Percent)
    ///             .set_background_color(Color::Yellow)
    ///             .set_foreground_color(Color::Red),
    ///     );
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill.png">
    ///
    pub fn set_pattern_fill(mut self, fill: &ShapePatternFill) -> ShapeFormat {
        self.pattern_fill = Some(fill.clone());
        self
    }

    /// Set the gradient fill formatting for a shape element.
    ///
    /// See the [`ShapeGradientFill`] struct for details on the gradient fill
    /// properties that can be set.
    ///
    /// # Parameters
    ///
    /// - `fill`: A [`ShapeGradientFill`] struct reference.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of its
    /// properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_format_set_gradient_fill.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeFormat, ShapeGradientFill, ShapeGradientStop, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox().set_text("This is some text").set_format(
    ///         &ShapeFormat::new().set_gradient_fill(&ShapeGradientFill::new().set_gradient_stops(&[
    ///             ShapeGradientStop::new("#F1DCDB", 0),
    ///             ShapeGradientStop::new("#963735", 100),
    ///         ])),
    ///     );
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_format_set_gradient_fill.png">
    ///
    pub fn set_gradient_fill(mut self, fill: &ShapeGradientFill) -> ShapeFormat {
        self.gradient_fill = Some(fill.clone());
        self
    }
}

// -----------------------------------------------------------------------
// ShapeLine
// -----------------------------------------------------------------------

/// The `ShapeLine` struct represents a shape line/border.
///
/// The [`ShapeLine`] struct represents the formatting properties for the line
/// of a Shape element. It is a sub property of the [`ShapeFormat`] struct and
/// is used with the [`ShapeFormat::set_line()`] method.
///
/// For 2D shapes the line property usually represents the border.
///
/// It is used in conjunction with the [`Shape`] struct.
///
/// # Examples
///
/// This example demonstrates adding a Textbox shape and setting some of the
/// line properties.
///
/// ```
/// # // This code is available in examples/doc_shape_line.rs
/// #
/// use rust_xlsxwriter::{Shape, ShapeLine, ShapeLineDashType, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Create a textbox shape with formatting.
///     let textbox = Shape::textbox().set_text("This is some text").set_format(
///         &ShapeLine::new()
///             .set_color("#FF0000")
///             .set_dash_type(ShapeLineDashType::DashDot),
///     );
///
///     // Insert a textbox in a cell.
///     worksheet.insert_shape(1, 1, &textbox)?;
///
///     // Save the file to disk.
///     workbook.save("shape.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_line.png">
///
#[derive(Clone, PartialEq)]
pub struct ShapeLine {
    pub(crate) color: Color,
    pub(crate) width: f64,
    pub(crate) transparency: u8,
    pub(crate) dash_type: ShapeLineDashType,
    pub(crate) hidden: bool,
}

impl ShapeLine {
    /// Create a new `ShapeLine` object to represent a Shape line/border.
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> ShapeLine {
        ShapeLine {
            color: Color::Default,
            width: 0.75,
            transparency: 0,
            dash_type: ShapeLineDashType::Solid,
            hidden: false,
        }
    }

    /// Set the color of a line/border.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a [`Color`] enum value or a
    ///   type that can convert [`Into`] a [`Color`] such as html like string.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// line properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_line_set_color.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeLine, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_format(&ShapeLine::new().set_color("#FF0000"));
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/shape_line_set_color.png">
    ///
    pub fn set_color(mut self, color: impl Into<Color>) -> ShapeLine {
        let color = color.into();
        if color.is_valid() {
            self.color = color;
        }

        self
    }

    /// Set the width of the line or border.
    ///
    /// # Parameters
    ///
    /// - `width`: The width should be specified in increments of 0.25 of a
    ///   point as in Excel. The width can be an number type that convert
    ///   [`Into`] [`f64`]. The default width is 0.75.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// line properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_line_set_width.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeLine, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_format(&ShapeLine::new().set_width(10));
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_set_width.png">
    ///
    pub fn set_width<T>(mut self, width: T) -> ShapeLine
    where
        T: Into<f64>,
    {
        let width = width.into();
        if width <= 1584.0 {
            self.width = width;
        }

        self
    }

    /// Set the dash type of the line or border.
    ///
    /// # Parameters
    ///
    /// - `dash_type`: A [`ShapeLineDashType`] enum value.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// line properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_line_set_dash_type.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeLine, ShapeLineDashType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_format(&ShapeLine::new().set_dash_type(ShapeLineDashType::DashDot));
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_set_dash_type.png">
    ///
    pub fn set_dash_type(mut self, dash_type: ShapeLineDashType) -> ShapeLine {
        self.dash_type = dash_type;
        self
    }

    /// Set the transparency of a line/border.
    ///
    /// Set the transparency of a line/border for a Shape element. You must also
    /// specify a line color in order for the transparency to be applied.
    ///
    /// # Parameters
    ///
    /// - `transparency`: The color transparency in the range 0 <= transparency
    ///   <= 100. The default value is 0.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// line properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_line_set_transparency.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeLine, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_format(&ShapeLine::new().set_color("#FF9900").set_transparency(50));
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_set_transparency.png">
    ///
    pub fn set_transparency(mut self, transparency: u8) -> ShapeLine {
        if transparency <= 100 {
            self.transparency = transparency;
        }

        self
    }

    /// Set the shape line as hidden.
    ///
    /// The method is sometimes required to turn off a default line type in
    /// order to highlight some other part of the line. This can also be
    /// achieved more succinctly using the [`ShapeFormat::set_no_line()`]
    /// method.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off (not hidden) by default.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// line properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_line_set_hidden.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeLine, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_format(&ShapeLine::new().set_hidden(true));
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/shape_line_set_hidden.png">
    ///
    pub fn set_hidden(mut self, enable: bool) -> ShapeLine {
        self.hidden = enable;
        self
    }
}

// -----------------------------------------------------------------------
// ShapeSolidFill
// -----------------------------------------------------------------------

/// The `ShapeSolidFill` struct represents a the solid fill for a shape element.
///
/// The [`ShapeSolidFill`] struct represents the formatting properties for the
/// solid fill of a Shape element. In Excel a solid fill is a single color fill
/// without a pattern or gradient.
///
/// `ShapeSolidFill` is a sub property of the [`ShapeFormat`] struct and is used
/// with the [`ShapeFormat::set_solid_fill()`] method.
///
/// It is used in conjunction with the [`Shape`] struct.
///
/// # Examples
///
/// This example demonstrates adding a Textbox shape and setting some of the
/// solid fill properties.
///
/// ```
/// # // This code is available in examples/doc_shape_solid_fill_set_transparency.rs
/// #
/// # use rust_xlsxwriter::{Shape, ShapeSolidFill, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet to the workbook.
/// #     let worksheet = workbook.add_worksheet();
/// #
///     // Create a textbox shape with formatting.
///     let textbox = Shape::textbox().set_text("This is some text").set_format(
///         &ShapeSolidFill::new()
///             .set_color("#8ED154")
///             .set_transparency(50),
///     );
///
///     // Insert a textbox in a cell.
///     worksheet.insert_shape(1, 1, &textbox)?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("shape.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_solid_fill_set_transparency.png">
///
#[derive(Clone, PartialEq)]
pub struct ShapeSolidFill {
    pub(crate) color: Color,
    pub(crate) transparency: u8,
}

impl ShapeSolidFill {
    /// Create a new `ShapeSolidFill` object to represent a Shape solid fill.
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> ShapeSolidFill {
        ShapeSolidFill {
            color: Color::Default,
            transparency: 0,
        }
    }

    /// Set the color of a solid fill.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a [`Color`] enum value or
    ///   a type that can convert [`Into`] a [`Color`].
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// solid fill properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_solid_fill_set_color.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeSolidFill, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_format(&ShapeSolidFill::new().set_color("#8ED154"));
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_solid_fill_set_color.png">
    ///
    pub fn set_color(mut self, color: impl Into<Color>) -> ShapeSolidFill {
        let color = color.into();
        if color.is_valid() {
            self.color = color;
        }

        self
    }

    /// Set the transparency of a solid fill.
    ///
    /// Set the transparency of a solid fill color for a Shape element. You must
    /// also specify a line color in order for the transparency to be applied.
    ///
    /// # Parameters
    ///
    /// - `transparency`: The color transparency in the range 0 <= transparency
    ///   <= 100. The default value is 0.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// solid fill properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_solid_fill_set_transparency.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeSolidFill, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox().set_text("This is some text").set_format(
    ///         &ShapeSolidFill::new()
    ///             .set_color("#8ED154")
    ///             .set_transparency(50),
    ///     );
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_solid_fill_set_transparency.png">
    ///
    pub fn set_transparency(mut self, transparency: u8) -> ShapeSolidFill {
        if transparency <= 100 {
            self.transparency = transparency;
        }

        self
    }
}

// -----------------------------------------------------------------------
// ShapePatternFill
// -----------------------------------------------------------------------

/// The `ShapePatternFill` struct represents a the pattern fill for a shape
/// element.
///
/// The [`ShapePatternFill`] struct represents the formatting properties for the
/// pattern fill of a Shape element. In Excel a pattern fill is comprised of a
/// simple pixelated pattern and background and foreground colors
///
/// `ShapePatternFill` is a sub property of the [`ShapeFormat`] struct and is
/// used with the [`ShapeFormat::set_pattern_fill()`] method.
///
/// It is used in conjunction with the [`Shape`] struct.
///
/// # Examples
///
/// This example demonstrates adding a Textbox shape and setting some of the
/// pattern fill properties.
///
/// ```
/// # // This code is available in examples/doc_shape_pattern_fill.rs
/// #
/// use rust_xlsxwriter::{Color, Shape, ShapePatternFill,
///                       ShapePatternFillType, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Create a textbox shape with formatting.
///     let textbox = Shape::textbox().set_text("This is some text").set_format(
///         &ShapePatternFill::new()
///             .set_pattern(ShapePatternFillType::Dotted20Percent)
///             .set_background_color(Color::Yellow)
///             .set_foreground_color(Color::Red),
///     );
///
///     // Insert a textbox in a cell.
///     worksheet.insert_shape(1, 1, &textbox)?;
///
///     // Save the file to disk.
///     workbook.save("shape.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill.png">
///
#[derive(Clone, PartialEq)]
pub struct ShapePatternFill {
    pub(crate) background_color: Color,
    pub(crate) foreground_color: Color,
    pub(crate) pattern: ShapePatternFillType,
}

impl ShapePatternFill {
    /// Create a new `ShapePatternFill` object to represent a Shape pattern fill.
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> ShapePatternFill {
        ShapePatternFill {
            background_color: Color::Default,
            foreground_color: Color::Default,
            pattern: ShapePatternFillType::Dotted5Percent,
        }
    }

    /// Set the pattern of a Shape pattern fill element.
    ///
    /// See the example above.
    ///
    /// # Parameters
    ///
    /// - `pattern`: The pattern property defined by a [`ShapePatternFillType`] enum value.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// pattern fill properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_pattern_fill.rs
    /// #
    /// # use rust_xlsxwriter::{Color, Shape, ShapePatternFill, ShapePatternFillType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox().set_text("This is some text").set_format(
    ///         &ShapePatternFill::new()
    ///             .set_pattern(ShapePatternFillType::Dotted20Percent)
    ///             .set_background_color(Color::Yellow)
    ///             .set_foreground_color(Color::Red),
    ///     );
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill.png">
    ///
    pub fn set_pattern(mut self, pattern: ShapePatternFillType) -> ShapePatternFill {
        self.pattern = pattern;
        self
    }

    /// Set the background color of a Shape pattern fill element.
    ///
    /// See the example above.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a [`Color`] enum value or
    ///   a type that can convert [`Into`] a [`Color`].
    ///
    pub fn set_background_color(mut self, color: impl Into<Color>) -> ShapePatternFill {
        let color = color.into();
        if color.is_valid() {
            self.background_color = color;
        }

        self
    }

    /// Set the foreground color of a Shape pattern fill element.
    ///
    /// See the example above.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a [`Color`] enum value or
    ///   a type that can convert [`Into`] a [`Color`].
    ///
    pub fn set_foreground_color(mut self, color: impl Into<Color>) -> ShapePatternFill {
        let color = color.into();
        if color.is_valid() {
            self.foreground_color = color;
        }

        self
    }
}

// -----------------------------------------------------------------------
// ShapeGradientFill
// -----------------------------------------------------------------------

/// The `ShapeGradientFill` struct represents a gradient fill for a shape
/// element.
///
/// The [`ShapeGradientFill`] struct represents the formatting properties for
/// the gradient fill of a Shape element. In Excel a gradient fill is comprised
/// of two or more colors that are blended gradually along a gradient.
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/gradient_fill_options.png">
///
/// `ShapeGradientFill` is a sub property of the [`ShapeFormat`] struct and is
/// used with the [`ShapeFormat::set_gradient_fill()`] method.
///
/// It is used in conjunction with the [`Shape`] struct.
///
/// # Examples
///
/// This example demonstrates adding a Textbox shape and setting some of the
/// gradient fill properties.
///
/// ```
/// # // This code is available in examples/doc_shape_gradient_fill.rs
/// #
/// use rust_xlsxwriter::{Shape, ShapeGradientFill, ShapeGradientStop, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///
///     // Add a worksheet to the workbook.
///     let worksheet = workbook.add_worksheet();
///
///     // Create a textbox shape with formatting.
///     let textbox = Shape::textbox().set_text("This is some text").set_format(
///         &ShapeGradientFill::new().set_gradient_stops(&[
///             ShapeGradientStop::new("#F1DCDB", 0),
///             ShapeGradientStop::new("#963735", 100),
///         ]),
///     );
///
///     // Insert a textbox in a cell.
///     worksheet.insert_shape(1, 1, &textbox)?;
///
///     // Save the file to disk.
///     workbook.save("shape.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_gradient_fill.png">
///
#[derive(Clone, PartialEq)]
pub struct ShapeGradientFill {
    pub(crate) gradient_type: ShapeGradientFillType,
    pub(crate) gradient_stops: Vec<ShapeGradientStop>,
    pub(crate) angle: u16,
}

impl Default for ShapeGradientFill {
    fn default() -> Self {
        Self::new()
    }
}

impl ShapeGradientFill {
    /// Create a new `ShapeGradientFill` object to represent a Shape gradient fill.
    ///
    pub fn new() -> ShapeGradientFill {
        ShapeGradientFill {
            gradient_type: ShapeGradientFillType::Linear,
            gradient_stops: vec![],
            angle: 90,
        }
    }

    /// Set the type of the gradient fill.
    ///
    /// Change the default type of the gradient fill to one of the styles
    /// supported by Excel.
    ///
    /// The four gradient types supported by Excel are:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_gradient_fill_types.png">
    ///
    /// # Parameters
    ///
    /// `gradient_type`: a [`ShapeGradientFillType`] enum value.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// gradient fill properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_gradient_fill_set_type.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeGradientFill, ShapeGradientFillType, ShapeGradientStop, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox().set_text("This is some text").set_format(
    ///         &ShapeGradientFill::new()
    ///             .set_type(ShapeGradientFillType::Rectangular)
    ///             .set_gradient_stops(&[
    ///                 ShapeGradientStop::new("#963735", 0),
    ///                 ShapeGradientStop::new("#F1DCDB", 100),
    ///             ]),
    ///     );
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_gradient_fill_set_type.png">
    ///
    pub fn set_type(mut self, gradient_type: ShapeGradientFillType) -> ShapeGradientFill {
        self.gradient_type = gradient_type;
        self
    }

    /// Set the gradient stops (data points) for a shape gradient fill.
    ///
    /// A gradient stop, encapsulated by the [`ShapeGradientStop`] struct,
    /// represents the properties of a data point that is used to generate a
    /// gradient fill.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/gradient_fill_options.png">
    ///
    /// Excel supports between 2 and 10 gradient stops which define the a color
    /// and its position in the gradient as a percentage. These colors and
    /// positions are used to interpolate a gradient fill.
    ///
    /// # Parameters
    ///
    /// `gradient_stops`: A slice ref of [`ShapeGradientStop`] values. As in
    /// Excel there must be between 2 and 10 valid gradient stops.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// gradient fill properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_gradient_fill_set_gradient_stops.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeGradientFill, ShapeGradientStop, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Set the properties of the gradient stops.
    ///     let gradient_stops = [
    ///         ShapeGradientStop::new("#DDEBCF", 0),
    ///         ShapeGradientStop::new("#9CB86E", 50),
    ///         ShapeGradientStop::new("#156B13", 100),
    ///     ];
    ///
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_format(&ShapeGradientFill::new().set_gradient_stops(&gradient_stops));
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_gradient_fill_set_gradient_stops.png">
    ///
    pub fn set_gradient_stops(mut self, gradient_stops: &[ShapeGradientStop]) -> ShapeGradientFill {
        let mut valid_gradient_stops = vec![];

        for gradient_stop in gradient_stops {
            if gradient_stop.is_valid() {
                valid_gradient_stops.push(gradient_stop.clone());
            }
        }

        if (2..=10).contains(&valid_gradient_stops.len()) {
            self.gradient_stops = valid_gradient_stops;
        } else {
            eprintln!("Gradient stops must contain between 2 and 10 valid entries.");
        }

        self
    }

    /// Set the angle of the linear gradient fill type.
    ///
    /// # Parameters
    ///
    /// - `angle`: The angle of the linear gradient fill in the range `0 <=
    ///   angle < 360`. The default angle is 90 degrees.
    ///
    pub fn set_angle(mut self, angle: u16) -> ShapeGradientFill {
        if (0..360).contains(&angle) {
            self.angle = angle;
        } else {
            eprintln!("Gradient angle '{angle}' must be in the Excel range 0 <= angle < 360");
        }
        self
    }
}

// -----------------------------------------------------------------------
// ShapeGradientStop
// -----------------------------------------------------------------------

/// The `ShapeGradientStop` struct represents a gradient fill data point.
///
/// The [`ShapeGradientStop`] struct represents the properties of a data point
/// (a stop) that is used to generate a gradient fill.
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/gradient_fill_options.png">
///
/// Excel supports between 2 and 10 gradient stops which define the a color and
/// its position in the gradient as a percentage. These colors and positions
/// are used to interpolate a gradient fill.
///
/// Gradient formats are generally used with the
/// [`ShapeGradientFill::set_gradient_stops()`] method and
/// [`ShapeGradientFill`].
///
/// # Examples
///
/// This example demonstrates adding a Textbox shape and setting some of the
/// gradient fill properties.
///
/// ```
/// # // This code is available in examples/doc_shape_gradient_fill_set_gradient_stops.rs
/// #
/// # use rust_xlsxwriter::{Shape, ShapeGradientFill, ShapeGradientStop, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet to the workbook.
/// #     let worksheet = workbook.add_worksheet();
///
///     // Set the properties of the gradient stops.
///     let gradient_stops = [
///         ShapeGradientStop::new("#DDEBCF", 0),
///         ShapeGradientStop::new("#9CB86E", 50),
///         ShapeGradientStop::new("#156B13", 100),
///     ];
///
///     // Create a textbox shape with formatting.
///     let textbox = Shape::textbox()
///         .set_text("This is some text")
///         .set_format(&ShapeGradientFill::new().set_gradient_stops(&gradient_stops));
///
///     // Insert a textbox in a cell.
///     worksheet.insert_shape(1, 1, &textbox)?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("shape.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_gradient_fill_set_gradient_stops.png">
///
#[derive(Clone, PartialEq)]
pub struct ShapeGradientStop {
    pub(crate) color: Color,
    pub(crate) position: u8,
}

impl ShapeGradientStop {
    /// Create a new `ShapeGradientStop` object to represent a Shape gradient fill stop.
    ///
    /// # Parameters
    ///
    /// - `color`: The gradient stop color property defined by a [`Color`] enum
    ///   value.
    /// - `position`: The gradient stop position in the range 0-100.
    ///
    /// # Examples
    ///
    /// An example of creating gradient stops for a gradient fill for a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_gradient_stops_new.rs
    /// #
    /// # use rust_xlsxwriter::ShapeGradientStop;
    /// #
    /// # #[allow(unused_variables)]
    /// # fn main() {
    ///     let gradient_stops = [
    ///         ShapeGradientStop::new("#156B13", 0),
    ///         ShapeGradientStop::new("#9CB86E", 50),
    ///         ShapeGradientStop::new("#DDEBCF", 100),
    ///     ];
    /// # }
    /// ```
    pub fn new(color: impl Into<Color>, position: u8) -> ShapeGradientStop {
        let color = color.into();

        // Check and warn but don't raise error since this is too deeply nested.
        // It will be rechecked and rejected at use.
        if !color.is_valid() {
            eprintln!("Gradient stop color isn't valid.");
        }
        if !(0..=100).contains(&position) {
            eprintln!("Gradient stop '{position}' outside Excel range: 0 <= position <= 100.");
        }

        ShapeGradientStop { color, position }
    }

    // Check for valid gradient stop properties.
    pub(crate) fn is_valid(&self) -> bool {
        self.color.is_valid() && (0..=100).contains(&self.position)
    }
}

// -----------------------------------------------------------------------
// ShapeText
// -----------------------------------------------------------------------

/// The `ShapeText` struct represents the text options for a shape element.
///
/// The [`ShapeText`] struct represents the text option properties for a Shape
///  element:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/shape_text_options_dialog.png">
///
/// Currently only the vertical, horizontal and text direction properties are
/// supported.
///
/// `ShapeText` is a sub property of the [`ShapeFormat`] struct and is used with
/// the [`Shape::set_text_options()`] method. See also [`ShapeFont`].
///
/// It is used in conjunction with the [`Shape`] struct.
///
/// # Examples
///
/// This example demonstrates adding a Textbox shape and setting some of the
/// text option properties.
///
///
/// ```
/// # // This code is available in examples/doc_shape_text_options_set_direction.rs
/// #
/// # use rust_xlsxwriter::{Shape, ShapeText, ShapeTextDirection, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet to the workbook.
/// #     let worksheet = workbook.add_worksheet();
///
///     // Create a textbox shape and add some text.
///     let textbox = Shape::textbox()
///         .set_text("古池や\n蛙飛び込む\n水の音")
///         .set_text_options(
///             &ShapeText::new().set_direction(ShapeTextDirection::Rotate90EastAsian)
///         );
///
///     // Insert a textbox in a cell.
///     worksheet.insert_shape(1, 1, &textbox)?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("shape.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_text_options_set_direction.png">
///
#[derive(Clone, PartialEq)]
pub struct ShapeText {
    pub(crate) horizontal_alignment: ShapeTextHorizontalAlignment,
    pub(crate) vertical_alignment: ShapeTextVerticalAlignment,
    pub(crate) direction: ShapeTextDirection,
}

impl Default for ShapeText {
    fn default() -> Self {
        Self::new()
    }
}

impl ShapeText {
    /// Create a new `ShapeText` object to represent the text options for a
    /// Shape element.
    ///
    pub fn new() -> ShapeText {
        ShapeText {
            horizontal_alignment: ShapeTextHorizontalAlignment::Default,
            vertical_alignment: ShapeTextVerticalAlignment::Top,
            direction: ShapeTextDirection::Horizontal,
        }
    }

    /// Set the horizontal alignment for the text in a shape textbox.
    ///
    /// This method sets the horizontal alignment for the text in a shape while
    /// [`ShapeText::set_vertical_alignment()`] sets the alignment for the text
    /// bounding box. See the example below.
    ///
    /// # Parameters
    ///
    /// - `alignment`: A [`ShapeTextHorizontalAlignment`] enum value.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// text option properties. This highlights the difference between
    /// horizontal and vertical centering.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_text_options_set_horizontal_alignment.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeText, ShapeTextHorizontalAlignment, ShapeTextVerticalAlignment, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Some text for the textbox.
    ///     let text = ["Some text", "on", "several lines"].join("\n");
    ///
    ///     // Create a textbox shape and add some text with horizontal alignment.
    ///     let textbox = Shape::textbox().set_text(&text).set_text_options(
    ///         &ShapeText::new().set_horizontal_alignment(ShapeTextHorizontalAlignment::Center),
    ///     );
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    ///
    ///     // Create a textbox shape and add some text with vertical alignment.
    ///     let textbox = Shape::textbox().set_text(&text).set_text_options(
    ///         &ShapeText::new().set_vertical_alignment(ShapeTextVerticalAlignment::TopCentered),
    ///     );
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 5, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/shape_text_set_horizontal_alignment.png">
    ///
    pub fn set_horizontal_alignment(
        mut self,
        alignment: ShapeTextHorizontalAlignment,
    ) -> ShapeText {
        self.horizontal_alignment = alignment;

        self
    }

    /// Set the vertical alignment for the textbox in a shape.
    ///
    /// This method sets the vertical alignment of the textbox in a shape while
    /// [`ShapeText::set_horizontal_alignment()`] sets the alignment for the
    /// text within the textbox. See the example above.
    ///
    /// # Parameters
    ///
    /// - `alignment`: A [`ShapeTextVerticalAlignment`] enum value.
    ///
    pub fn set_vertical_alignment(mut self, alignment: ShapeTextVerticalAlignment) -> ShapeText {
        self.vertical_alignment = alignment;

        self
    }

    /// Set the text direction of the text in the text box.
    ///
    /// This is useful for languages that display text vertically.
    ///
    /// # Parameters
    ///
    /// - `direction`: The [`ShapeTextDirection`] of the text.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// text option properties.
    ///
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_text_options_set_direction.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeText, ShapeTextDirection, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape and add some text.
    ///     let textbox = Shape::textbox()
    ///         .set_text("古池や\n蛙飛び込む\n水の音")
    ///         .set_text_options(
    ///             &ShapeText::new().set_direction(ShapeTextDirection::Rotate90EastAsian)
    ///         );
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_options_set_direction.png">
    ///
    pub fn set_direction(mut self, direction: ShapeTextDirection) -> ShapeText {
        self.direction = direction;

        self
    }
}

// -----------------------------------------------------------------------
// Shape enums
// -----------------------------------------------------------------------

// The `ShapeType` enum defines the [`Shape`] types.
//
// Note, currently only the type supported is TextBox. See the explanation at
// the start of this document.
//
#[derive(Clone, PartialEq, Eq)]
pub(crate) enum ShapeType {
    // The Textbox shape.
    TextBox,
}

/// The `ShapeLineDashType` enum defines the [`Shape`] line dash types.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ShapeLineDashType {
    /// Solid - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_solid.png">
    Solid,

    /// Round dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_round_dot.png">
    RoundDot,

    /// Square dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_square_dot.png">
    SquareDot,

    /// Dash - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_dash.png">
    Dash,

    /// Dash dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_dash_dot.png">
    DashDot,

    /// Long dash - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash.png">
    LongDash,

    /// Long dash dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash_dot.png">
    LongDashDot,

    /// Long dash dot dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_line_dash_longdash_dot_dot.png">
    LongDashDotDot,
}

impl fmt::Display for ShapeLineDashType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Dash => write!(f, "dash"),
            Self::Solid => write!(f, "solid"),
            Self::DashDot => write!(f, "dashDot"),
            Self::LongDash => write!(f, "lgDash"),
            Self::RoundDot => write!(f, "sysDot"),
            Self::SquareDot => write!(f, "sysDash"),
            Self::LongDashDot => write!(f, "lgDashDot"),
            Self::LongDashDotDot => write!(f, "lgDashDotDot"),
        }
    }
}

/// The `ShapePatternFillType` enum defines the [`Shape`] pattern fill types.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ShapePatternFillType {
    /// Dotted 5 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_5_percent.png">
    Dotted5Percent,

    /// Dotted 10 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_10_percent.png">
    Dotted10Percent,

    /// Dotted 20 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_20_percent.png">
    Dotted20Percent,

    /// Dotted 25 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_25_percent.png">
    Dotted25Percent,

    /// Dotted 30 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_30_percent.png">
    Dotted30Percent,

    /// Dotted 40 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_40_percent.png">
    Dotted40Percent,

    /// Dotted 50 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_50_percent.png">
    Dotted50Percent,

    /// Dotted 60 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_60_percent.png">
    Dotted60Percent,

    /// Dotted 70 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_70_percent.png">
    Dotted70Percent,

    /// Dotted 75 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_75_percent.png">
    Dotted75Percent,

    /// Dotted 80 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_80_percent.png">
    Dotted80Percent,

    /// Dotted 90 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_90_percent.png">
    Dotted90Percent,

    /// Diagonal stripes light downwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_light_downwards.png">
    DiagonalStripesLightDownwards,

    /// Diagonal stripes light upwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_light_upwards.png">
    DiagonalStripesLightUpwards,

    /// Diagonal stripes dark downwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_dark_downwards.png">
    DiagonalStripesDarkDownwards,

    /// Diagonal stripes dark upwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_dark_upwards.png">
    DiagonalStripesDarkUpwards,

    /// Diagonal stripes wide downwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_wide_downwards.png">
    DiagonalStripesWideDownwards,

    /// Diagonal stripes wide upwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_stripes_wide_upwards.png">
    DiagonalStripesWideUpwards,

    /// Vertical stripes light - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_light.png">
    VerticalStripesLight,

    /// Horizontal stripes light - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_light.png">
    HorizontalStripesLight,

    /// Vertical stripes narrow - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_narrow.png">
    VerticalStripesNarrow,

    /// Horizontal stripes narrow - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_narrow.png">
    HorizontalStripesNarrow,

    /// Vertical stripes dark - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_dark.png">
    VerticalStripesDark,

    /// Horizontal stripes dark - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_dark.png">
    HorizontalStripesDark,

    /// Stripes backslashes - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_stripes_backslashes.png">
    StripesBackslashes,

    /// Stripes forward slashes - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_stripes_forward_slashes.png">
    StripesForwardSlashes,

    /// Horizontal stripes alternating - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_stripes_alternating.png">
    HorizontalStripesAlternating,

    /// Vertical stripes alternating - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_vertical_stripes_alternating.png">
    VerticalStripesAlternating,

    /// Small confetti - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_small_confetti.png">
    SmallConfetti,

    /// Large confetti - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_large_confetti.png">
    LargeConfetti,

    /// Zigzag - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_zigzag.png">
    Zigzag,

    /// Wave - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_wave.png">
    Wave,

    /// Diagonal brick - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_diagonal_brick.png">
    DiagonalBrick,

    /// Horizontal brick - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_horizontal_brick.png">
    HorizontalBrick,

    /// Weave - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_weave.png">
    Weave,

    /// Plaid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_plaid.png">
    Plaid,

    /// Divot - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_divot.png">
    Divot,

    /// Dotted grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_grid.png">
    DottedGrid,

    /// Dotted diamond - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_dotted_diamond.png">
    DottedDiamond,

    /// Shingle - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_shingle.png">
    Shingle,

    /// Trellis - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_trellis.png">
    Trellis,

    /// Sphere - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_sphere.png">
    Sphere,

    /// Small grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_small_grid.png">
    SmallGrid,

    /// Large grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_large_grid.png">
    LargeGrid,

    /// Small checkerboard - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_small_checkerboard.png">
    SmallCheckerboard,

    /// Large checkerboard - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_large_checkerboard.png">
    LargeCheckerboard,

    /// Outlined diamond grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_outlined_diamond_grid.png">
    OutlinedDiamondGrid,

    /// Solid diamond grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/chart_pattern_fill_solid_diamond_grid.png">
    SolidDiamondGrid,
}

impl fmt::Display for ShapePatternFillType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Wave => write!(f, "wave"),
            Self::Weave => write!(f, "weave"),
            Self::Plaid => write!(f, "plaid"),
            Self::Divot => write!(f, "divot"),
            Self::Zigzag => write!(f, "zigZag"),
            Self::Sphere => write!(f, "sphere"),
            Self::Shingle => write!(f, "shingle"),
            Self::Trellis => write!(f, "trellis"),
            Self::SmallGrid => write!(f, "smGrid"),
            Self::LargeGrid => write!(f, "lgGrid"),
            Self::DottedGrid => write!(f, "dotGrid"),
            Self::DottedDiamond => write!(f, "dotDmnd"),
            Self::DiagonalBrick => write!(f, "diagBrick"),
            Self::LargeConfetti => write!(f, "lgConfetti"),
            Self::SmallConfetti => write!(f, "smConfetti"),
            Self::Dotted5Percent => write!(f, "pct5"),
            Self::Dotted10Percent => write!(f, "pct10"),
            Self::Dotted20Percent => write!(f, "pct20"),
            Self::Dotted25Percent => write!(f, "pct25"),
            Self::Dotted30Percent => write!(f, "pct30"),
            Self::Dotted40Percent => write!(f, "pct40"),
            Self::Dotted50Percent => write!(f, "pct50"),
            Self::Dotted60Percent => write!(f, "pct60"),
            Self::Dotted70Percent => write!(f, "pct70"),
            Self::Dotted75Percent => write!(f, "pct75"),
            Self::Dotted80Percent => write!(f, "pct80"),
            Self::Dotted90Percent => write!(f, "pct90"),
            Self::HorizontalBrick => write!(f, "horzBrick"),
            Self::SolidDiamondGrid => write!(f, "solidDmnd"),
            Self::SmallCheckerboard => write!(f, "smCheck"),
            Self::LargeCheckerboard => write!(f, "lgCheck"),
            Self::StripesBackslashes => write!(f, "dashDnDiag"),
            Self::VerticalStripesDark => write!(f, "dkVert"),
            Self::OutlinedDiamondGrid => write!(f, "openDmnd"),
            Self::VerticalStripesLight => write!(f, "ltVert"),
            Self::HorizontalStripesDark => write!(f, "dkHorz"),
            Self::StripesForwardSlashes => write!(f, "dashUpDiag"),
            Self::VerticalStripesNarrow => write!(f, "narVert"),
            Self::HorizontalStripesLight => write!(f, "ltHorz"),
            Self::HorizontalStripesNarrow => write!(f, "narHorz"),
            Self::DiagonalStripesDarkUpwards => write!(f, "dkUpDiag"),
            Self::DiagonalStripesWideUpwards => write!(f, "wdUpDiag"),
            Self::VerticalStripesAlternating => write!(f, "dashVert"),
            Self::DiagonalStripesLightUpwards => write!(f, "ltUpDiag"),
            Self::DiagonalStripesDarkDownwards => write!(f, "dkDnDiag"),
            Self::DiagonalStripesWideDownwards => write!(f, "wdDnDiag"),
            Self::HorizontalStripesAlternating => write!(f, "dashHorz"),
            Self::DiagonalStripesLightDownwards => write!(f, "ltDnDiag"),
        }
    }
}

/// The `ShapeTextHorizontalAlignment` enum defines the horizontal alignment for
/// [`Shape`] text.
///
/// See [`ShapeText::set_horizontal_alignment()`].
#[derive(Clone, PartialEq, Eq, Default)]
pub enum ShapeTextHorizontalAlignment {
    /// Horizontally align text in the default position (usually to the left).
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_horizontal_alignment_default.png">
    #[default]
    Default,

    /// Horizontally align text to the left of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_horizontal_alignment_default.png">
    Left,

    /// Horizontally align text to the center of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_horizontal_alignment_center.png">
    Center,

    /// Horizontally align text to the right of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_horizontal_alignment_right.png">
    Right,
}

/// The `ShapeTextVerticalAlignment` enum defines the vertical alignment for
/// [`Shape`] text.
///
/// See [`ShapeText::set_horizontal_alignment()`].
#[derive(Clone, PartialEq, Eq, Default)]
pub enum ShapeTextVerticalAlignment {
    /// Vertically align text to the top of the shape. This is the default.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_horizontal_alignment_default.png">
    #[default]
    Top,

    /// Vertically align text to the middle of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_vertical_alignment_middle.png">
    Middle,

    /// Vertically align text to the bottom of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_vertical_alignment_bottom.png">
    Bottom,

    /// Vertically align text to the top center of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_vertical_alignment_top_centered.png">
    TopCentered,

    /// Vertically align text to the middle center of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_vertical_alignment_middle_centered.png">
    MiddleCentered,

    /// Vertically align text to the bottom center of the shape.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_vertical_alignment_bottom_centered.png">
    BottomCentered,
}

/// The `ShapeTextDirection` enum defines the text direction for [`Shape`] text.
///
/// See [`ShapeText::set_direction()`].
#[derive(Clone, PartialEq, Eq, Default)]
pub enum ShapeTextDirection {
    /// Text is horizontal. This is the Excel default.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_direction_horizontal.png">
    #[default]
    Horizontal,

    /// Text is rotated 270 degrees.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_direction_rotate_270.png">
    Rotate270,

    /// Text is rotated 90 degrees.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_direction_rotate_90.png">
    Rotate90,

    /// Text direction is rotated 90 degrees but the characters aren't rotated. Suitable for East Asian text.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_direction_rotate_90_east_asian.png">
    Rotate90EastAsian,

    /// Text is stacked vertically.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_text_direction_stacked.png">
    Stacked,
}

// -----------------------------------------------------------------------
// ShapeFont
// -----------------------------------------------------------------------

#[derive(Clone, PartialEq)]
/// The `ShapeFont` struct represents the font format for shape objects.
///
/// Excel uses a standard font dialog for text elements of a shape such as the
/// shape title or axes data labels. It looks like this:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_font_dialog.png">
///
/// The [`ShapeFont`] struct represents many of these font options such as font
/// type, size, color and properties such as bold and italic. It is used in
/// conjunction with the [`Shape::set_font()`] method.
///
/// It is used in conjunction with the [`Shape`] struct.
///
/// # Examples
///
/// This example demonstrates adding a Textbox shape and setting some of the
/// font properties.
///
/// ```
/// # // This code is available in examples/doc_shape_set_font.rs
/// #
/// # use rust_xlsxwriter::{Shape, ShapeFont, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet to the workbook.
/// #     let worksheet = workbook.add_worksheet();
/// #
///     // Create a textbox shape with formatting.
///     let textbox = Shape::textbox().set_text("This is some text").set_font(
///         &ShapeFont::new()
///             .set_bold()
///             .set_italic()
///             .set_name("American Typewriter")
///             .set_color("#0000FF")
///             .set_size(15),
///     );
///
///     // Insert a textbox in a cell.
///     worksheet.insert_shape(1, 1, &textbox)?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("shape.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_set_font.png">
///
pub struct ShapeFont {
    // Shape/axis titles have a default bold font so we need to handle that as
    // an option.
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    pub(crate) underline: bool,
    pub(crate) name: String,
    pub(crate) size: f64,
    pub(crate) color: Color,
    pub(crate) strikethrough: bool,
    pub(crate) pitch_family: u8,
    pub(crate) character_set: u8,
    pub(crate) rotation: Option<i16>,
    pub(crate) has_baseline: bool,
    pub(crate) right_to_left: Option<bool>,
}

impl Default for ShapeFont {
    fn default() -> Self {
        Self::new()
    }
}

impl ShapeFont {
    /// Create a new `ShapeFont` object to represent a Shape font.
    ///
    pub fn new() -> ShapeFont {
        ShapeFont {
            bold: false,
            italic: false,
            underline: false,
            name: String::new(),
            size: 1100.0,
            color: Color::Default,
            strikethrough: false,
            pitch_family: 0,
            character_set: 0,
            rotation: None,
            has_baseline: false,
            right_to_left: None,
        }
    }

    /// Set the bold property for the font of a shape element.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// font properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_font_set_bold.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFont, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_font(&ShapeFont::new().set_bold());
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_font_set_bold.png">
    ///
    pub fn set_bold(mut self) -> ShapeFont {
        self.bold = true;
        self
    }

    /// Set the italic property for the font of a shape element.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// font properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_font_set_italic.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFont, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_font(&ShapeFont::new().set_italic());
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_font_set_italic.png">
    ///
    pub fn set_italic(mut self) -> ShapeFont {
        self.italic = true;
        self
    }

    /// Set the color property for the font of a shape element.
    ///
    /// # Parameters
    ///
    /// - `color`: The font color property defined by a [`Color`] enum
    ///   value.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// font properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_font_set_color.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFont, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_font(&ShapeFont::new().set_color("#FF0000"));
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_font_set_color.png">
    ///
    pub fn set_color(mut self, color: impl Into<Color>) -> ShapeFont {
        let color = color.into();
        if color.is_valid() {
            self.color = color;
        }

        self
    }

    /// Set the shape font name property.
    ///
    /// Set the name/type of a font for a shape element.
    ///
    /// # Parameters
    ///
    /// - `font_name`: The font name property.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// font properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_font_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFont, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_font(&ShapeFont::new().set_name("American Typewriter"));
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_font_set_name.png">
    ///
    pub fn set_name(mut self, font_name: impl Into<String>) -> ShapeFont {
        self.name = font_name.into();
        self
    }

    /// Set the size property for the font of a shape element.
    ///
    /// # Parameters
    ///
    /// - `font_size`: The font size property.
    ///
    /// # Examples
    ///
    /// This example demonstrates adding a Textbox shape and setting some of the
    /// font properties.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_font_set_size.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFont, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet to the workbook.
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    ///     // Create a textbox shape with formatting.
    ///     let textbox = Shape::textbox()
    ///         .set_text("This is some text")
    ///         .set_font(&ShapeFont::new().set_size(20));
    ///
    ///     // Insert a textbox in a cell.
    ///     worksheet.insert_shape(1, 1, &textbox)?;
    /// #
    /// #     // Save the file to disk.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_font_set_size.png">
    ///
    pub fn set_size<T>(mut self, font_size: T) -> ShapeFont
    where
        T: Into<f64>,
    {
        self.size = font_size.into() * 100.0;
        self
    }

    /// Set the underline property for the font of a shape element.
    ///
    /// The default underline type is the only type supported.
    ///
    pub fn set_underline(mut self) -> ShapeFont {
        self.underline = true;
        self
    }

    /// Set the strikethrough property for the font of a shape element.
    ///
    pub fn set_strikethrough(mut self) -> ShapeFont {
        self.strikethrough = true;
        self
    }

    /// Unset the bold property for a font.
    ///
    /// Some shape elements such as titles have a default bold property in
    /// Excel. This method can be used to turn it off.
    ///
    pub fn unset_bold(mut self) -> ShapeFont {
        self.bold = false;
        self
    }

    /// Display the shape font from right to left for some language support.
    ///
    /// See
    /// [`Worksheet::set_right_to_left()`](crate::Worksheet::set_right_to_left)
    /// for details.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_right_to_left(mut self, enable: bool) -> ShapeFont {
        self.right_to_left = Some(enable);
        self
    }

    /// Set the pitch family property for the font of a shape element.
    ///
    /// This function is implemented for completeness but is rarely used in
    /// practice.
    ///
    /// # Parameters
    ///
    /// - `family`: The font family property.
    ///
    pub fn set_pitch_family(mut self, family: u8) -> ShapeFont {
        self.pitch_family = family;
        self
    }

    /// Set the character set property for the font of a shape element.
    ///
    /// Set the font character set. This function is implemented for
    /// completeness but is rarely required in practice.
    ///
    /// # Parameters
    ///
    /// - `character_set`: The font character set property.
    ///
    pub fn set_character_set(mut self, character_set: u8) -> ShapeFont {
        self.character_set = character_set;
        self
    }

    // Internal check for font properties that need a sub-element.
    pub(crate) fn is_latin(&self) -> bool {
        !self.name.is_empty() || self.pitch_family > 0 || self.character_set > 0
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
/// The `ShapeGradientFillType` enum defines the gradient types of a
/// [`ShapeGradientFill`].
///
/// The four gradient types supported by Excel are:
///
/// <img src="https://rustxlsxwriter.github.io/images/chart_gradient_fill_types.png">
///
pub enum ShapeGradientFillType {
    /// The gradient runs linearly from the top of the area vertically to the
    /// bottom. This is the default.
    Linear,

    /// The gradient runs radially from the bottom right of the area vertically
    /// to the top left.
    Radial,

    /// The gradient runs in a rectangular pattern from the bottom right of the
    /// area vertically to the top left.
    Rectangular,

    /// The gradient runs in a rectangular pattern from the center of the area
    /// to the outer vertices.
    Path,
}

// -----------------------------------------------------------------------
// Traits
// -----------------------------------------------------------------------

// Trait for objects that have a component stored in the drawing.xml file.
impl DrawingObject for Shape {
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

// -----------------------------------------------------------------------
// IntoShapeFormat
// -----------------------------------------------------------------------

/// Trait to map types into a `ShapeFormat`.
///
/// The `IntoShapeFormat` trait provides a syntactic shortcut for the
/// [`Shape::set_format()`] method to allow it to handle sub-structs such as:
///
/// - [`ShapeLine`]
/// - [`ShapeSolidFill`]
/// - [`ShapeGradientFill`]
/// - [`ShapePatternFill`]
///
/// In order to pass one of these sub-structs as a parameter you would normally
/// have to create a [`ShapeFormat`] and then add the sub-struct, as shown in
/// the first part of the example below. However, this can be a little verbose
/// if you just want to format one of the sub-properties. The `IntoShapeFormat`
/// trait will accept the sub-structs listed above and create a parent
/// [`ShapeFormat`] instance to wrap it in, see the second part of the example
/// below.
///
/// # Examples
///
/// An example of passing shape formatting parameters via the
/// [`IntoShapeFormat`] trait.
///
/// ```
/// # // This code is available in examples/doc_into_shape_format.rs
/// #
/// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeSolidFill, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet to the workbook.
/// #     let worksheet = workbook.add_worksheet();
///
///     // Create a formatted shape via ShapeFormat and ShapeSolidFill.
///     let textbox = Shape::textbox()
///         .set_text("This is some text").set_format(
///             &ShapeFormat::new().set_solid_fill(&ShapeSolidFill::new().set_color("#8ED154")),
///     );
///
///     // Insert a textbox in a cell.
///     worksheet.insert_shape(1, 1, &textbox)?;
///
///     // Create a formatted shape via ShapeSolidFill directly.
///     let textbox = Shape::textbox()
///         .set_text("This is some text")
///         .set_format(&ShapeSolidFill::new().set_color("#8ED154"));
///
///     // Insert a textbox in a cell.
///     worksheet.insert_shape(1, 5, &textbox)?;
/// #
/// #     // Save the file to disk.
/// #     workbook.save("shape.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/into_shape_format.png">
///
pub trait IntoShapeFormat {
    /// Trait function to turn a type into [`ShapeFormat`].
    fn new_shape_format(&self) -> ShapeFormat;
}

impl IntoShapeFormat for &ShapeFormat {
    fn new_shape_format(&self) -> ShapeFormat {
        (*self).clone()
    }
}

impl IntoShapeFormat for &ShapeLine {
    fn new_shape_format(&self) -> ShapeFormat {
        ShapeFormat::new().set_line(self).clone()
    }
}

impl IntoShapeFormat for &ShapeSolidFill {
    fn new_shape_format(&self) -> ShapeFormat {
        ShapeFormat::new().set_solid_fill(self).clone()
    }
}

impl IntoShapeFormat for &ShapePatternFill {
    fn new_shape_format(&self) -> ShapeFormat {
        ShapeFormat::new().set_pattern_fill(self).clone()
    }
}

impl IntoShapeFormat for &ShapeGradientFill {
    fn new_shape_format(&self) -> ShapeFormat {
        ShapeFormat::new().set_gradient_fill(self).clone()
    }
}
