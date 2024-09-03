// shape - A module to represent Excel cell shapes.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::fmt;

use crate::drawing::{DrawingObject, DrawingType};
use crate::{Color, ObjectMovement, Url};

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
    pub(crate) format: ShapeFormat,
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
            format: ShapeFormat::default(),
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

    /// Set the [`ShapeFormat`] of the shape.
    ///
    /// Set the font or background properties of a shape using a [`ShapeFormat`]
    /// object. TODO
    ///
    /// This API is currently experimental and may go away in the future.
    ///
    /// # Parameters
    ///
    /// - `format`: The [`ShapeFormat`] property for the shape.
    ///
    pub fn set_format<T>(mut self, format: T) -> Shape
    where
        T: IntoShapeFormat,
    {
        self.format = format.new_shape_format();
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

// -----------------------------------------------------------------------
// ShapeFormat
// -----------------------------------------------------------------------

#[derive(Clone, PartialEq)]
/// The `ShapeFormat` struct represents formatting for various shape objects.
///
/// Excel uses a standard formatting dialog for the elements of a shape such as
/// data series, the plot area, the shape area, the legend or individual points.
/// It looks like this:
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
/// - [`ShapeFormat::set_border()`]: Set the [`ShapeBorder`] properties. A
///   synonym for [`ShapeLine`] depending on context.
/// - [`ShapeFormat::set_no_line()`]: Turn off the line for the shape object.
/// - [`ShapeFormat::set_no_border()`]: Turn off the border for the shape
///   object.
///
/// # Examples
///
/// An example of accessing the [`ShapeFormat`] for data series in a shape and
/// using them to apply formatting.
///
/// ```
/// # // This code is available in examples/app_shape_pattern.rs
/// #
/// # use rust_xlsxwriter::*;
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #     let bold = Format::new().set_bold();
/// #
/// #     // Add the worksheet data that the shapes will refer to.
/// #     worksheet.write_with_format(0, 0, "Shingle", &bold)?;
/// #     worksheet.write_with_format(0, 1, "Brick", &bold)?;
/// #
/// #     let data = [[105, 150, 130, 90], [50, 120, 100, 110]];
/// #     for (col_num, col_data) in data.iter().enumerate() {
/// #         for (row_num, row_data) in col_data.iter().enumerate() {
/// #             worksheet.write(row_num as u32 + 1, col_num as u16, *row_data)?;
/// #         }
/// #     }
/// #
/// #     // Create a new column shape.
///     let mut shape = Shape::new(ShapeType::Column);
///
///     // Configure the first data series and add fill patterns.
///     shape
///         .add_series()
///         .set_name("Sheet1!$A$1")
///         .set_values("Sheet1!$A$2:$A$5")
///         .set_gap(70)
///         .set_format(
///             ShapeFormat::new()
///                 .set_pattern_fill(
///                     ShapePatternFill::new()
///                         .set_pattern(ShapePatternFillType::Shingle)
///                         .set_foreground_color("#804000")
///                         .set_background_color("#C68C53"),
///                 )
///                 .set_border(ShapeLine::new().set_color("#804000")),
///         );
///
///     shape
///         .add_series()
///         .set_name("Sheet1!$B$1")
///         .set_values("Sheet1!$B$2:$B$5")
///         .set_format(
///             ShapeFormat::new()
///                 .set_pattern_fill(
///                     ShapePatternFill::new()
///                         .set_pattern(ShapePatternFillType::HorizontalBrick)
///                         .set_foreground_color("#B30000")
///                         .set_background_color("#FF6666"),
///                 )
///                 .set_border(ShapeLine::new().set_color("#B30000")),
///         );
///
///     // Add a shape title and some axis labels.
///     shape.title().set_name("Cladding types");
///     shape.x_axis().set_name("Region");
///     shape.y_axis().set_name("Number of houses");
///
///     // Add the shape to the worksheet.
///     worksheet.insert_shape(1, 3, &shape)?;
///
///     workbook.save("shape_pattern.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/app_shape_pattern.png">
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
    pub fn set_line(&mut self, line: &ShapeLine) -> &mut ShapeFormat {
        self.line = Some(line.clone());
        self
    }

    /// Set the border formatting for a shape element.
    ///
    /// See the [`ShapeLine`] struct for details on the border properties that
    /// can be set. As a syntactic shortcut you can use the type alias
    /// [`ShapeBorder`] instead
    /// of `ShapeLine`.
    ///
    /// # Parameters
    ///
    /// - `line`: A [`ShapeLine`] struct reference.
    ///
    /// # Examples
    ///
    /// An example of formatting the border in a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_border_formatting.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeBorder, ShapeFormat, ShapeLineDashType, ShapeType, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ShapeFormat::new()
    ///                 .set_border(
    ///                     ShapeBorder::new()
    ///                         .set_color("#FF9900")
    ///                         .set_width(5.25)
    ///                         .set_dash_type(ShapeLineDashType::SquareDot)
    ///                         .set_transparency(70),
    ///                 )
    ///                 .set_no_fill(),
    ///         );
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_border_formatting.png">
    ///
    pub fn set_border(&mut self, line: &ShapeLine) -> &mut ShapeFormat {
        self.set_line(line)
    }

    /// Turn off the line property for a shape element.
    ///
    /// The line property for a shape element can be turned off if you wish to
    /// hide it.
    ///
    /// # Examples
    ///
    /// An example of turning off a default line in a shape format.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_format_set_no_line.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 1)?;
    /// #     worksheet.write(1, 0, 2)?;
    /// #     worksheet.write(2, 0, 3)?;
    /// #     worksheet.write(3, 0, 4)?;
    /// #     worksheet.write(4, 0, 5)?;
    /// #     worksheet.write(5, 0, 6)?;
    /// #     worksheet.write(0, 1, 10)?;
    /// #     worksheet.write(1, 1, 40)?;
    /// #     worksheet.write(2, 1, 50)?;
    /// #     worksheet.write(3, 1, 20)?;
    /// #     worksheet.write(4, 1, 10)?;
    /// #     worksheet.write(5, 1, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::ScatterStraightWithMarkers);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_categories("Sheet1!$A$1:$A$6")
    ///         .set_values("Sheet1!$B$1:$B$6")
    ///         .set_format(ShapeFormat::new().set_no_line());
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_format_set_no_line.png">
    ///
    pub fn set_no_line(&mut self) -> &mut ShapeFormat {
        self.no_line = true;
        self
    }

    /// Turn off the border property for a shape element.
    ///
    /// The border property for a shape element can be turned off if you wish to
    /// hide it.
    ///
    /// # Examples
    ///
    /// An example of turning off the border of a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_format_set_no_border.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ShapeFormat::new().set_no_border());
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_format_set_no_border.png">
    ///
    pub fn set_no_border(&mut self) -> &mut ShapeFormat {
        self.set_no_line()
    }

    /// Turn off the fill property for a shape element.
    ///
    /// The fill property for a shape element can be turned off if you wish to
    /// hide it and display only the border (if set).
    ///
    /// # Examples
    ///
    /// An example of turning off the fill of a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_format_set_no_fill.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeLine, ShapeType, Workbook, Color, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ShapeFormat::new()
    ///                 .set_border(ShapeLine::new().set_color(Color::Black))
    ///                 .set_no_fill(),
    ///         );
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
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
    pub fn set_no_fill(&mut self) -> &mut ShapeFormat {
        self.no_fill = true;
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
    /// An example of setting a solid fill for a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_solid_fill.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeSolidFill, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ShapeFormat::new().set_solid_fill(
    ///                 ShapeSolidFill::new()
    ///                     .set_color("#FF9900")
    ///                     .set_transparency(60),
    ///             ),
    ///         );
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_solid_fill.png">
    ///
    pub fn set_solid_fill(&mut self, fill: &ShapeSolidFill) -> &mut ShapeFormat {
        self.solid_fill = Some(fill.clone());
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
    /// An example of setting a pattern fill for a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_pattern_fill.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeFormat, ShapePatternFill, ShapePatternFillType, ShapeType, Workbook, Color,
    /// #     XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ShapeFormat::new().set_pattern_fill(
    ///                 ShapePatternFill::new()
    ///                     .set_pattern(ShapePatternFillType::Dotted20Percent)
    ///                     .set_background_color(Color::Yellow)
    ///                     .set_foreground_color(Color::Red),
    ///             ),
    ///         );
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
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
    pub fn set_pattern_fill(&mut self, fill: &ShapePatternFill) -> &mut ShapeFormat {
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
    /// An example of setting a gradient fill for a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_format_set_gradient_fill.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeFormat, ShapeGradientFill, ShapeGradientStop, ShapeType, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ShapeFormat::new().set_gradient_fill(
    ///             ShapeGradientFill::new().set_gradient_stops(&[
    ///                 ShapeGradientStop::new("#963735", 0),
    ///                 ShapeGradientStop::new("#F1DCDB", 100),
    ///             ]),
    ///         ));
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_gradient_fill.png">
    ///
    pub fn set_gradient_fill(&mut self, fill: &ShapeGradientFill) -> &mut ShapeFormat {
        self.gradient_fill = Some(fill.clone());
        self
    }
}

/// The `ShapeLine` struct represents a shape line/border.
///
/// The [`ShapeLine`] struct represents the formatting properties for a line or
/// border for a Shape element. It is a sub property of the [`ShapeFormat`]
/// struct and is used with the [`ShapeFormat::set_line()`] or
/// [`ShapeFormat::set_border()`] methods.
///
/// Excel uses the element names "Line" and "Border" depending on the context.
/// For a Line shape the line is represented by a line property but for a Column
/// shape the line becomes the border. Both of these share the same properties
/// and are both represented in `rust_xlsxwriter` by the [`ShapeLine`] struct.
///
/// As a syntactic shortcut you can use the type alias [`ShapeBorder`] instead
/// of `ShapeLine`.
///
/// It is used in conjunction with the [`Shape`] struct.
///
/// # Examples
///
/// An example of formatting a line/border in a shape element.
///
/// ```
/// # // This code is available in examples/doc_shape_line_formatting.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Shape, ShapeFormat, ShapeLine, ShapeLineDashType, ShapeType, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the shape.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(3, 0, 20)?;
/// #     worksheet.write(4, 0, 10)?;
/// #     worksheet.write(5, 0, 50)?;
/// #
/// #     // Create a new shape.
///     let mut shape = Shape::new(ShapeType::Line);
///
///     // Add a data series with formatting.
///     shape
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_format(
///             ShapeFormat::new().set_line(
///                 ShapeLine::new()
///                     .set_color("#FF9900")
///                     .set_width(5.25)
///                     .set_dash_type(ShapeLineDashType::SquareDot)
///                     .set_transparency(70),
///             ),
///         );
///
///     // Add the shape to the worksheet.
///     worksheet.insert_shape(0, 2, &shape)?;
///
/// #     // Save the file.
/// #     workbook.save("shape.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/shape_line_formatting.png">
///
#[derive(Clone, PartialEq)]
pub struct ShapeLine {
    pub(crate) color: Color,
    pub(crate) width: Option<f64>,
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
            width: None,
            transparency: 0,
            dash_type: ShapeLineDashType::Solid,
            hidden: false,
        }
    }

    /// Set the color of a line/border.
    ///
    /// # Parameters
    ///
    /// - `color`: The color property defined by a [`Color`] enum value or
    ///   a type that can convert [`Into`] a [`Color`].
    ///
    /// # Examples
    ///
    /// An example of formatting the line color in a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_line_set_color.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeLine, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Line);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ShapeFormat::new().set_line(ShapeLine::new().set_color("#FF9900")));
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_set_color.png">
    ///
    pub fn set_color(&mut self, color: impl Into<Color>) -> &mut ShapeLine {
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
    ///   [`Into`] [`f64`].
    ///
    /// # Examples
    ///
    /// An example of formatting the line width in a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_line_set_width.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeLine, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Line);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ShapeFormat::new().set_line(ShapeLine::new().set_width(10.0)));
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/shape_line_set_width.png">
    ///
    pub fn set_width<T>(&mut self, width: T) -> &mut ShapeLine
    where
        T: Into<f64>,
    {
        let width = width.into();
        if width <= 1584.0 {
            self.width = Some(width);
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
    /// An example of formatting the line dash type in a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_line_set_dash_type.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeFormat, ShapeLine, ShapeLineDashType, ShapeType, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Line);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ShapeFormat::new()
    ///                 .set_line(ShapeLine::new()
    ///                 .set_dash_type(ShapeLineDashType::DashDot)),
    ///         );
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
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
    pub fn set_dash_type(&mut self, dash_type: ShapeLineDashType) -> &mut ShapeLine {
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
    /// An example of formatting the line transparency in a shape element. Note, you
    /// must set also set a color in order to set the transparency.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_line_set_transparency.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeLine, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Line);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ShapeFormat::new().set_line(ShapeLine::new().set_color("#FF9900").set_transparency(50)),
    ///         );
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
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
    pub fn set_transparency(&mut self, transparency: u8) -> &mut ShapeLine {
        if transparency <= 100 {
            self.transparency = transparency;
        }

        self
    }

    /// Set the shape line as hidden.
    ///
    /// The method is sometimes required to turn off a default line type in
    /// order to highlight some other element such as the line markers.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off (not hidden) by
    ///   default.
    ///
    pub fn set_hidden(&mut self, enable: bool) -> &mut ShapeLine {
        self.hidden = enable;
        self
    }
}

/// A type to represent a Shape border. It can be used interchangeably with
/// [`ShapeLine`].
///
/// Excel uses the shape element names "Line" and "Border" depending on the
/// context. For a Line shape the line is represented by a line property but for
/// a Column shape the line becomes the border. Both of these share the same
/// properties and are both represented in `rust_xlsxwriter` by the
/// [`ShapeLine`] struct.
///
/// The `ShapeBorder` type is a type alias of [`ShapeLine`] for use as a
/// syntactic shortcut where you would expect to write `ShapeBorder` instead of
/// `ShapeLine`.
///
/// # Examples
///
/// An example of formatting the border in a shape element.
///
/// ```
/// # // This code is available in examples/doc_shape_border_formatting.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Shape, ShapeBorder, ShapeFormat, ShapeLineDashType, ShapeType, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the shape.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(3, 0, 20)?;
/// #     worksheet.write(4, 0, 10)?;
/// #     worksheet.write(5, 0, 50)?;
/// #
/// #     // Create a new shape.
///     let mut shape = Shape::new(ShapeType::Column);
///
///     // Add a data series with formatting.
///     shape
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_format(
///             ShapeFormat::new()
///                 .set_border(
///                     ShapeBorder::new()
///                         .set_color("#FF9900")
///                         .set_width(5.25)
///                         .set_dash_type(ShapeLineDashType::SquareDot)
///                         .set_transparency(70),
///                 )
///                 .set_no_fill(),
///         );
///
///     // Add the shape to the worksheet.
///     worksheet.insert_shape(0, 2, &shape)?;
///
/// #     // Save the file.
/// #     workbook.save("shape.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/shape_border_formatting.png">
///
pub type ShapeBorder = ShapeLine;

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
/// An example of setting a solid fill for a shape element.
///
/// ```
/// # // This code is available in examples/doc_shape_solid_fill.rs
/// #
/// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeSolidFill, ShapeType, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the shape.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(3, 0, 20)?;
/// #     worksheet.write(4, 0, 10)?;
/// #     worksheet.write(5, 0, 50)?;
/// #
/// #     // Create a new shape.
///     let mut shape = Shape::new(ShapeType::Column);
///
///     // Add a data series with formatting.
///     shape
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_format(
///             ShapeFormat::new().set_solid_fill(
///                 ShapeSolidFill::new()
///                     .set_color("#FF9900")
///                     .set_transparency(60),
///             ),
///         );
///
///     // Add the shape to the worksheet.
///     worksheet.insert_shape(0, 2, &shape)?;
///
/// #     // Save the file.
/// #     workbook.save("shape.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_solid_fill.png">
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
    /// An example of setting a solid fill color for a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_solid_fill_set_color.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeSolidFill, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ShapeFormat::new().set_solid_fill(ShapeSolidFill::new().set_color("#B5A401")));
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
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
    pub fn set_color(&mut self, color: impl Into<Color>) -> &mut ShapeSolidFill {
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
    /// An example of setting a solid fill for a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_solid_fill.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeSolidFill, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ShapeFormat::new().set_solid_fill(
    ///                 ShapeSolidFill::new()
    ///                     .set_color("#FF9900")
    ///                     .set_transparency(60),
    ///             ),
    ///         );
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_solid_fill.png">
    ///
    pub fn set_transparency(&mut self, transparency: u8) -> &mut ShapeSolidFill {
        if transparency <= 100 {
            self.transparency = transparency;
        }

        self
    }
}

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
/// An example of setting a pattern fill for a shape element.
///
/// ```
/// # // This code is available in examples/doc_shape_pattern_fill.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Shape, ShapeFormat, ShapePatternFill, ShapePatternFillType, ShapeType, Workbook, Color, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the shape.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(3, 0, 20)?;
/// #     worksheet.write(4, 0, 10)?;
/// #     worksheet.write(5, 0, 50)?;
/// #
/// #     // Create a new shape.
///     let mut shape = Shape::new(ShapeType::Column);
///
///     // Add a data series with formatting.
///     shape
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_format(
///             ShapeFormat::new().set_pattern_fill(
///                 ShapePatternFill::new()
///                     .set_pattern(ShapePatternFillType::Dotted20Percent)
///                     .set_background_color(Color::Yellow)
///                     .set_foreground_color(Color::Red),
///             ),
///         );
///
///     // Add the shape to the worksheet.
///     worksheet.insert_shape(0, 2, &shape)?;
///
/// #     // Save the file.
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
    ///
    /// # Examples
    ///
    /// An example of setting a pattern fill for a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_pattern_fill_set_pattern.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeFormat, ShapePatternFill, ShapePatternFillType, ShapeType, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ShapeFormat::new().set_pattern_fill(
    ///             ShapePatternFill::new().set_pattern(ShapePatternFillType::DiagonalBrick),
    ///         ));
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_set_pattern.png">
    ///
    pub fn set_pattern(&mut self, pattern: ShapePatternFillType) -> &mut ShapePatternFill {
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
    /// # Examples
    ///
    /// An example of setting a pattern fill for a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_pattern_fill.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeFormat, ShapePatternFill, ShapePatternFillType, ShapeType, Workbook, Color,
    /// #     XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ShapeFormat::new().set_pattern_fill(
    ///                 ShapePatternFill::new()
    ///                     .set_pattern(ShapePatternFillType::Dotted20Percent)
    ///                     .set_background_color(Color::Yellow)
    ///                     .set_foreground_color(Color::Red),
    ///             ),
    ///         );
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
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
    pub fn set_background_color(&mut self, color: impl Into<Color>) -> &mut ShapePatternFill {
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
    pub fn set_foreground_color(&mut self, color: impl Into<Color>) -> &mut ShapePatternFill {
        let color = color.into();
        if color.is_valid() {
            self.foreground_color = color;
        }

        self
    }
}

/// The `ShapeLineDashType` enum defines the [`Shape`] line dash types.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ShapeLineDashType {
    /// Solid - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_dash_solid.png">
    Solid,

    /// Round dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_dash_round_dot.png">
    RoundDot,

    /// Square dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_dash_square_dot.png">
    SquareDot,

    /// Dash - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_dash_dash.png">
    Dash,

    /// Dash dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_dash_dash_dot.png">
    DashDot,

    /// Long dash - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_dash_longdash.png">
    LongDash,

    /// Long dash dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_dash_longdash_dot.png">
    LongDashDot,

    /// Long dash dot dot - shape line/border dash type.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_line_dash_longdash_dot_dot.png">
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
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_5_percent.png">
    Dotted5Percent,

    /// Dotted 10 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_10_percent.png">
    Dotted10Percent,

    /// Dotted 20 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_20_percent.png">
    Dotted20Percent,

    /// Dotted 25 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_25_percent.png">
    Dotted25Percent,

    /// Dotted 30 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_30_percent.png">
    Dotted30Percent,

    /// Dotted 40 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_40_percent.png">
    Dotted40Percent,

    /// Dotted 50 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_50_percent.png">
    Dotted50Percent,

    /// Dotted 60 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_60_percent.png">
    Dotted60Percent,

    /// Dotted 70 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_70_percent.png">
    Dotted70Percent,

    /// Dotted 75 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_75_percent.png">
    Dotted75Percent,

    /// Dotted 80 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_80_percent.png">
    Dotted80Percent,

    /// Dotted 90 percent - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_90_percent.png">
    Dotted90Percent,

    /// Diagonal stripes light downwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_diagonal_stripes_light_downwards.png">
    DiagonalStripesLightDownwards,

    /// Diagonal stripes light upwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_diagonal_stripes_light_upwards.png">
    DiagonalStripesLightUpwards,

    /// Diagonal stripes dark downwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_diagonal_stripes_dark_downwards.png">
    DiagonalStripesDarkDownwards,

    /// Diagonal stripes dark upwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_diagonal_stripes_dark_upwards.png">
    DiagonalStripesDarkUpwards,

    /// Diagonal stripes wide downwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_diagonal_stripes_wide_downwards.png">
    DiagonalStripesWideDownwards,

    /// Diagonal stripes wide upwards - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_diagonal_stripes_wide_upwards.png">
    DiagonalStripesWideUpwards,

    /// Vertical stripes light - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_vertical_stripes_light.png">
    VerticalStripesLight,

    /// Horizontal stripes light - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_horizontal_stripes_light.png">
    HorizontalStripesLight,

    /// Vertical stripes narrow - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_vertical_stripes_narrow.png">
    VerticalStripesNarrow,

    /// Horizontal stripes narrow - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_horizontal_stripes_narrow.png">
    HorizontalStripesNarrow,

    /// Vertical stripes dark - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_vertical_stripes_dark.png">
    VerticalStripesDark,

    /// Horizontal stripes dark - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_horizontal_stripes_dark.png">
    HorizontalStripesDark,

    /// Stripes backslashes - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_stripes_backslashes.png">
    StripesBackslashes,

    /// Stripes forward slashes - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_stripes_forward_slashes.png">
    StripesForwardSlashes,

    /// Horizontal stripes alternating - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_horizontal_stripes_alternating.png">
    HorizontalStripesAlternating,

    /// Vertical stripes alternating - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_vertical_stripes_alternating.png">
    VerticalStripesAlternating,

    /// Small confetti - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_small_confetti.png">
    SmallConfetti,

    /// Large confetti - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_large_confetti.png">
    LargeConfetti,

    /// Zigzag - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_zigzag.png">
    Zigzag,

    /// Wave - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_wave.png">
    Wave,

    /// Diagonal brick - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_diagonal_brick.png">
    DiagonalBrick,

    /// Horizontal brick - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_horizontal_brick.png">
    HorizontalBrick,

    /// Weave - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_weave.png">
    Weave,

    /// Plaid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_plaid.png">
    Plaid,

    /// Divot - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_divot.png">
    Divot,

    /// Dotted grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_grid.png">
    DottedGrid,

    /// Dotted diamond - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_dotted_diamond.png">
    DottedDiamond,

    /// Shingle - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_shingle.png">
    Shingle,

    /// Trellis - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_trellis.png">
    Trellis,

    /// Sphere - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_sphere.png">
    Sphere,

    /// Small grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_small_grid.png">
    SmallGrid,

    /// Large grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_large_grid.png">
    LargeGrid,

    /// Small checkerboard - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_small_checkerboard.png">
    SmallCheckerboard,

    /// Large checkerboard - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_large_checkerboard.png">
    LargeCheckerboard,

    /// Outlined diamond grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_outlined_diamond_grid.png">
    OutlinedDiamondGrid,

    /// Solid diamond grid - shape fill pattern.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/shape_pattern_fill_solid_diamond_grid.png">
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

// -----------------------------------------------------------------------
// ShapeFont
// -----------------------------------------------------------------------

#[derive(Clone, PartialEq)]
/// The `ShapeFont` struct represents the font format for various shape objects.
///
/// Excel uses a standard font dialog for text elements of a shape such as the
/// shape title or axes data labels. It looks like this:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_font_dialog.png">
///
/// The [`ShapeFont`] struct represents many of these font options such as font
/// type, size, color and properties such as bold and italic. It is generally
/// used in conjunction with a `set_font()` method for a shape element.
///
/// It is used in conjunction with the [`Shape`] struct.
///
/// # Examples
///
/// An example of setting the font for a shape element.
///
/// ```
/// # // This code is available in examples/doc_shape_font.rs
/// #
/// # use rust_xlsxwriter::{Shape, ShapeFont, ShapeType, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the shape.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(3, 0, 20)?;
/// #     worksheet.write(4, 0, 10)?;
/// #     worksheet.write(5, 0, 50)?;
/// #
/// #     // Create a new shape.
///     let mut shape = Shape::new(ShapeType::Column);
///
///     // Add a data series.
///     shape.add_series().set_values("Sheet1!$A$1:$A$6");
///
///     // Set the font for an axis.
///     shape.x_axis().set_font(
///         ShapeFont::new()
///             .set_bold()
///             .set_italic()
///             .set_name("Consolas")
///             .set_color("#FF0000"),
///     );
///
///     // Hide legend for clarity.
///     shape.legend().set_hidden();
///
///     // Add the shape to the worksheet.
///     worksheet.insert_shape(0, 2, &shape)?;
///
/// #     // Save the file.
/// #     workbook.save("shape.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_font.png">
///
pub struct ShapeFont {
    // Shape/axis titles have a default bold font so we need to handle that as
    // an option.
    pub(crate) bold: Option<bool>,
    pub(crate) has_default_bold: bool,

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
            bold: None,
            italic: false,
            underline: false,
            name: String::new(),
            size: 0.0,
            color: Color::Default,
            strikethrough: false,
            pitch_family: 0,
            character_set: 0,
            rotation: None,
            has_baseline: false,
            has_default_bold: false,
            right_to_left: None,
        }
    }

    /// Set the bold property for the font of a shape element.
    ///
    /// # Examples
    ///
    /// An example of setting the bold property for the font in a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_font_set_bold.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFont, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series.
    ///     shape.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     shape.x_axis().set_font(ShapeFont::new().set_bold());
    ///
    ///     // Hide legend for clarity.
    ///     shape.legend().set_hidden();
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
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
    pub fn set_bold(&mut self) -> &mut ShapeFont {
        self.bold = Some(true);
        self
    }

    /// Set the italic property for the font of a shape element.
    ///
    /// # Examples
    ///
    /// An example of setting the italic property for the font in a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_font_set_italic.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFont, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series.
    ///     shape.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     shape.x_axis().set_font(ShapeFont::new().set_italic());
    ///
    ///     // Hide legend for clarity.
    ///     shape.legend().set_hidden();
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
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
    pub fn set_italic(&mut self) -> &mut ShapeFont {
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
    /// An example of setting the color property for the font in a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_font_set_color.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFont, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series.
    ///     shape.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     shape.x_axis().set_font(ShapeFont::new().set_color("#FF0000"));
    ///
    ///     // Hide legend for clarity.
    ///     shape.legend().set_hidden();
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
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
    pub fn set_color(&mut self, color: impl Into<Color>) -> &mut ShapeFont {
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
    ///
    /// # Examples
    ///
    /// An example of setting the font name property for the font in a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_font_set_name.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFont, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series.
    ///     shape.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     shape
    ///         .x_axis()
    ///         .set_font(ShapeFont::new().set_name("American Typewriter"));
    ///
    ///     // Hide legend for clarity.
    ///     shape.legend().set_hidden();
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
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
    pub fn set_name(&mut self, font_name: impl Into<String>) -> &mut ShapeFont {
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
    /// An example of setting the font size property for the font in a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_font_set_size.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFont, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series.
    ///     shape.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     shape.x_axis().set_font(ShapeFont::new().set_size(20));
    ///
    ///     // Hide legend for clarity.
    ///     shape.legend().set_hidden();
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
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
    pub fn set_size<T>(&mut self, font_size: T) -> &mut ShapeFont
    where
        T: Into<f64>,
    {
        self.size = font_size.into() * 100.0;
        self
    }

    /// Set the text rotation property for the font of a shape element.
    ///
    /// Set the rotation angle of the text in a cell. The rotation can be any
    /// angle in the range -90 to 90 degrees, or 270-271 to indicate text where
    /// the letters run from top to bottom, see below.
    ///
    /// # Parameters
    ///
    /// - `rotation`: The rotation angle in the range `-90 <= rotation <= 90`.
    ///   Two special case values are supported:
    ///   - `270`: Stacked text, where the text runs from top to bottom.
    ///   - `271`: A special variant of stacked text for East Asian fonts.
    ///
    /// # Examples
    ///
    /// An example of setting the font text rotation for the font in a shape
    /// element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_font_set_rotation.rs
    /// #
    /// # use rust_xlsxwriter::{Shape, ShapeFont, ShapeType, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series.
    ///     shape.add_series().set_values("Sheet1!$A$1:$A$6");
    ///
    ///     // Set the font for an axis.
    ///     shape.x_axis().set_font(ShapeFont::new().set_rotation(45));
    ///
    ///     // Hide legend for clarity.
    ///     shape.legend().set_hidden();
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/shape_font_set_rotation.png">
    ///
    pub fn set_rotation(&mut self, rotation: i16) -> &mut ShapeFont {
        match rotation {
            270..=271 | -90..=90 => self.rotation = Some(rotation),
            _ => eprintln!("Rotation '{rotation}' outside range: -90 <= angle <= 90."),
        }

        self
    }

    /// Set the underline property for the font of a shape element.
    ///
    /// The default underline type is the only type supported.
    ///
    pub fn set_underline(&mut self) -> &mut ShapeFont {
        self.underline = true;
        self
    }

    /// Set the strikethrough property for the font of a shape element.
    ///
    pub fn set_strikethrough(&mut self) -> &mut ShapeFont {
        self.strikethrough = true;
        self
    }

    /// Unset the bold property for a font.
    ///
    /// Some shape elements such as titles have a default bold property in
    /// Excel. This method can be used to turn it off.
    ///
    pub fn unset_bold(&mut self) -> &mut ShapeFont {
        self.bold = Some(false);
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
    pub fn set_right_to_left(&mut self, enable: bool) -> &mut ShapeFont {
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
    pub fn set_pitch_family(&mut self, family: u8) -> &mut ShapeFont {
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
    pub fn set_character_set(&mut self, character_set: u8) -> &mut ShapeFont {
        self.character_set = character_set;
        self
    }

    #[doc(hidden)]
    /// Set the default bold property for the font.
    ///
    /// The is mainly only required for testing to ensure strict compliance with
    /// Excel's output.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is off by default.
    ///
    pub fn set_default_bold(&mut self, enable: bool) -> &mut ShapeFont {
        self.has_default_bold = enable;
        self
    }
}

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
///
/// # Examples
///
/// An example of setting a gradient fill for a shape element.
///
/// ```
/// # // This code is available in examples/doc_shape_gradient_fill.rs
/// #
/// use rust_xlsxwriter::{
///     Shape, ShapeGradientFill, ShapeGradientStop, ShapeType, Workbook, XlsxError,
/// };
///
/// fn main() -> Result<(), XlsxError> {
///     let mut workbook = Workbook::new();
///     let worksheet = workbook.add_worksheet();
///
///     // Add some data for the shape.
///     worksheet.write(0, 0, 10)?;
///     worksheet.write(1, 0, 40)?;
///     worksheet.write(2, 0, 50)?;
///     worksheet.write(3, 0, 20)?;
///     worksheet.write(4, 0, 10)?;
///     worksheet.write(5, 0, 50)?;
///
///     // Create a new shape.
///   let mut shape = Shape::new(ShapeType::Column);
///
///     // Add a data series with formatting.
///     shape
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_format(ShapeGradientFill::new().set_gradient_stops(&[
///             ShapeGradientStop::new("#963735", 0),
///             ShapeGradientStop::new("#F1DCDB", 100),
///         ]));
///
///     // Add the shape to the worksheet.
///     worksheet.insert_shape(0, 2, &shape)?;
///
///     // Save the file.
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

// -----------------------------------------------------------------------
// ShapeGradientFill
// -----------------------------------------------------------------------

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
    /// <img src="https://rustxlsxwriter.github.io/images/shape_gradient_fill_types.png">
    ///
    /// # Parameters
    ///
    /// `gradient_type`: a [`ShapeGradientFillType`] enum value.
    ///
    /// # Examples
    ///
    /// An example of setting a gradient fill for a shape element with a non-default
    /// gradient type.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_gradient_fill_set_type.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeGradientFill, ShapeGradientFillType, ShapeGradientStop, ShapeType, Workbook,
    /// #     XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(
    ///             ShapeGradientFill::new()
    ///                 .set_type(ShapeGradientFillType::Rectangular)
    ///                 .set_gradient_stops(&[
    ///                     ShapeGradientStop::new("#963735", 0),
    ///                     ShapeGradientStop::new("#F1DCDB", 100),
    ///                 ]),
    ///         );
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
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
    pub fn set_type(&mut self, gradient_type: ShapeGradientFillType) -> &mut ShapeGradientFill {
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
    /// An example of setting a gradient fill for a shape element.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_gradient_stops.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeGradientFill, ShapeGradientStop, ShapeType, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Set the properties of the gradient stops.
    ///     let gradient_stops = [
    ///         ShapeGradientStop::new("#156B13", 0),
    ///         ShapeGradientStop::new("#9CB86E", 50),
    ///         ShapeGradientStop::new("#DDEBCF", 100),
    ///     ];
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ShapeGradientFill::new().set_gradient_stops(&gradient_stops));
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/shape_gradient_stops.png">
    ///
    /// Note, it can be clearer to add the gradient stops directly to the format
    /// as follows. This gives the same output as above.
    ///
    /// ```
    /// # // This code is available in examples/doc_shape_gradient_stops2.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     Shape, ShapeGradientFill, ShapeGradientStop, ShapeType, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some data for the shape.
    /// #     worksheet.write(0, 0, 10)?;
    /// #     worksheet.write(1, 0, 40)?;
    /// #     worksheet.write(2, 0, 50)?;
    /// #     worksheet.write(3, 0, 20)?;
    /// #     worksheet.write(4, 0, 10)?;
    /// #     worksheet.write(5, 0, 50)?;
    /// #
    /// #     // Create a new shape.
    ///     let mut shape = Shape::new(ShapeType::Column);
    ///
    ///     // Add a data series with formatting.
    ///     shape
    ///         .add_series()
    ///         .set_values("Sheet1!$A$1:$A$6")
    ///         .set_format(ShapeGradientFill::new().set_gradient_stops(&[
    ///             ShapeGradientStop::new("#156B13", 0),
    ///             ShapeGradientStop::new("#9CB86E", 50),
    ///             ShapeGradientStop::new("#DDEBCF", 100),
    ///         ]));
    ///
    ///     // Add the shape to the worksheet.
    ///     worksheet.insert_shape(0, 2, &shape)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("shape.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    pub fn set_gradient_stops(
        &mut self,
        gradient_stops: &[ShapeGradientStop],
    ) -> &mut ShapeGradientFill {
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
    pub fn set_angle(&mut self, angle: u16) -> &mut ShapeGradientFill {
        if (0..360).contains(&angle) {
            self.angle = angle;
        } else {
            eprintln!("Gradient angle '{angle}' must be in the Excel range 0 <= angle < 360");
        }
        self
    }
}

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
/// An example of setting a gradient fill for a shape element.
///
/// ```
/// # // This code is available in examples/doc_shape_gradient_stops.rs
/// #
/// # use rust_xlsxwriter::{
/// #     Shape, ShapeGradientFill, ShapeGradientStop, ShapeType, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the shape.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(3, 0, 20)?;
/// #     worksheet.write(4, 0, 10)?;
/// #     worksheet.write(5, 0, 50)?;
/// #
/// #     // Create a new shape.
///     let mut shape = Shape::new(ShapeType::Column);
///
///     // Set the properties of the gradient stops.
///     let gradient_stops = [
///         ShapeGradientStop::new("#156B13", 0),
///         ShapeGradientStop::new("#9CB86E", 50),
///         ShapeGradientStop::new("#DDEBCF", 100),
///     ];
///
///     // Add a data series with formatting.
///     shape
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$6")
///         .set_format(ShapeGradientFill::new().set_gradient_stops(&gradient_stops));
///
///     // Add the shape to the worksheet.
///     worksheet.insert_shape(0, 2, &shape)?;
///
/// #     // Save the file.
/// #     workbook.save("shape.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_gradient_stops.png">
///
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

#[derive(Clone, Copy, PartialEq, Eq)]
/// The `ShapeGradientFillType` enum defines the gradient types of a
/// [`ShapeGradientFill`].
///
/// The four gradient types supported by Excel are:
///
/// <img src="https://rustxlsxwriter.github.io/images/shape_gradient_fill_types.png">
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
// IntoShapeFormat
// -----------------------------------------------------------------------

/// Trait to map types into a `ShapeFormat`.
///
/// The `IntoShapeFormat` trait provides a syntactic shortcut for the
/// `shape.*.set_format()` methods that take [`ShapeFormat`] as a parameter.
///
/// The [`ShapeFormat`] struct mirrors the Excel Shape element formatting dialog
/// and has several sub-structs such as:
///
/// - [`ShapeLine`]
/// - [`ShapeSolidFill`]
/// - [`ShapePatternFill`]
///
/// In order to pass one of these sub-structs as a parameter you would normally
/// have to create a [`ShapeFormat`] first and then add the sub-struct, as shown
/// in the first part of the example below. However, since this is a little
/// verbose if you just want to format one of the sub-properties the
/// `IntoShapeFormat` trait will accept the sub-structs listed above and create
/// a parent [`ShapeFormat`] instance to wrap it in, see the second part of the
/// example below.
///
/// # Examples
///
/// An example of passing shape formatting parameters via the
/// [`IntoShapeFormat`] trait
///
/// ```
/// # // This code is available in examples/doc_into_shape_format.rs
/// #
/// # use rust_xlsxwriter::{Shape, ShapeFormat, ShapeSolidFill, ShapeType, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some data for the shape.
/// #     worksheet.write(0, 0, 10)?;
/// #     worksheet.write(1, 0, 40)?;
/// #     worksheet.write(2, 0, 50)?;
/// #     worksheet.write(0, 1, 20)?;
/// #     worksheet.write(1, 1, 10)?;
/// #     worksheet.write(2, 1, 50)?;
/// #
/// #     // Create a new shape.
///     let mut shape = Shape::new(ShapeType::Column);
///
///     // Add formatting via ShapeFormat and a ShapeSolidFill sub struct.
///     shape
///         .add_series()
///         .set_values("Sheet1!$A$1:$A$3")
///         .set_format(ShapeFormat::new().set_solid_fill(ShapeSolidFill::new().set_color("#40EABB")));
///
///     // Add formatting using a ShapeSolidFill struct directly.
///     shape
///         .add_series()
///         .set_values("Sheet1!$B$1:$B$3")
///         .set_format(ShapeSolidFill::new().set_color("#AAC3F2"));
///
///     // Add the shape to the worksheet.
///     worksheet.insert_shape(0, 2, &shape)?;
///
/// #     // Save the file.
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

impl IntoShapeFormat for &mut ShapeFormat {
    fn new_shape_format(&self) -> ShapeFormat {
        (*self).clone()
    }
}

impl IntoShapeFormat for &mut ShapeLine {
    fn new_shape_format(&self) -> ShapeFormat {
        ShapeFormat::new().set_line(self).clone()
    }
}

impl IntoShapeFormat for &mut ShapeSolidFill {
    fn new_shape_format(&self) -> ShapeFormat {
        ShapeFormat::new().set_solid_fill(self).clone()
    }
}

impl IntoShapeFormat for &mut ShapePatternFill {
    fn new_shape_format(&self) -> ShapeFormat {
        ShapeFormat::new().set_pattern_fill(self).clone()
    }
}

impl IntoShapeFormat for &mut ShapeGradientFill {
    fn new_shape_format(&self) -> ShapeFormat {
        ShapeFormat::new().set_gradient_fill(self).clone()
    }
}
