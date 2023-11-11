// image - A module to represent Excel conditional formats.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! # Working with Conditional Formats
//!
//! Conditional formatting is a feature of Excel which allows you to apply a
//! format to a cell or a range of cells based on certain criteria. For example
//! you might apply rules like the following to highlight cells in different
//! ranges.
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/conditional_format_dialog.png">
//!
//! With `rust_xlsxwriter` we could emulate this by creating [`Format`]
//! instances and applying them to conditional format rules, like this:
//!
//! ```
//! # // This code is available in examples/doc_conditional_format_cell1.rs
//! #
//! # use rust_xlsxwriter::{
//! #     ConditionalFormatCell, ConditionalFormatCellCriteria, Format, Workbook, XlsxError,
//! # };
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     // Create a new Excel file object.
//! #     let mut workbook = Workbook::new();
//! #     let worksheet = workbook.add_worksheet();
//! #
//! #     // Add some sample data.
//! #     let data = [
//! #         [90, 80, 50, 10, 20, 90, 40, 90, 30, 40],
//! #         [20, 10, 90, 100, 30, 60, 70, 60, 50, 90],
//! #         [10, 50, 60, 50, 20, 50, 80, 30, 40, 60],
//! #         [10, 90, 20, 40, 10, 40, 50, 70, 90, 50],
//! #         [70, 100, 10, 90, 10, 10, 20, 100, 100, 40],
//! #         [20, 60, 10, 100, 30, 10, 20, 60, 100, 10],
//! #         [10, 60, 10, 80, 100, 80, 30, 30, 70, 40],
//! #         [30, 90, 60, 10, 10, 100, 40, 40, 30, 40],
//! #         [80, 90, 10, 20, 20, 50, 80, 20, 60, 90],
//! #         [60, 80, 30, 30, 10, 50, 80, 60, 50, 30],
//! #     ];
//! #     worksheet.write_row_matrix(2, 1, data)?;
//! #
//! #     // Set the column widths for clarity.
//! #     for col_num in 1..=10u16 {
//! #         worksheet.set_column_width(col_num, 6)?;
//! #     }
//! #
//!     // Add a format. Light red fill with dark red text.
//!     let format1 = Format::new()
//!         .set_font_color("9C0006")
//!         .set_background_color("FFC7CE");
//!
//!     // Add a format. Green fill with dark green text.
//!     let format2 = Format::new()
//!         .set_font_color("006100")
//!         .set_background_color("C6EFCE");
//!
//!     // Write a conditional format over a range.
//!     let conditional_format = ConditionalFormatCell::new()
//!         .set_criteria(ConditionalFormatCellCriteria::GreaterThanOrEqualTo)
//!         .set_value(50)
//!         .set_format(format1);
//!
//!     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
//!
//!     // Write another conditional format over the same range.
//!     let conditional_format = ConditionalFormatCell::new()
//!         .set_criteria(ConditionalFormatCellCriteria::LessThan)
//!         .set_value(50)
//!         .set_format(format2);
//!
//!     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
//!
//! #     // Save the file.
//! #     workbook.save("conditional_format.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Which would produce an output like the following. Cells with values >= 50
//! are in light red. Values < 50 are in light green.
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/conditional_format_cell1.png">
//!
//! # Replicating an Excel conditional format with `rust_xlsxwriter`
//!
//! It is important not to try to reverse engineer Excel's conditional
//! formatting rules from `rust_xlsxwriter`. If you aren't familiar with the
//! syntax and functionality of conditional formats then a better place to start
//! is in Excel. Create a conditional format in Excel to meet your needs and
//! then port it over to `rust_xlsxwriter`.
//!
//! There are several common features of all conditional formats:
//!
//! - A range: The range that the conditional format applies to. This is usually
//!   set via the
//!   [`Worksheet::add_conditional_format()`](crate::Worksheet::add_conditional_format)
//!   method.
//! - A rule: This can be an equality like `>=` or a rule like "Top 10".
//! - A target: This is usually a cell or range that the rule applies to. This
//!   mainly applies to Cell style conditional formats. For other types of
//!   conditional format the "range" is the target.
//! - A format: The cell format with properties such as text or background color
//!   to high the cell if the rule matches.
//!
//! The following are the structs that represent the main conditional format
//! variants in Excel. See each of these sections for more information:
//!
//! - [`ConditionalFormatCell`]: The Cell style conditional format. This is the
//!   most common style of conditional formats which uses simple equalities such
//!   as "equal to" or "greater than" or "between". See the example above.
//! - [`ConditionalFormatAverage`]: The Average/Standard Deviation style
//!   conditional format.
//! - [`ConditionalFormatDuplicate`]: The Duplicate/Unique style conditional
//!   format.
//! - [`ConditionalFormatText`]: The Text conditional format for rules like
//!   "contains" or "begins with".
//! - [`ConditionalFormatTop`]: The Top/Bottom style conditional format.
//!
//! # Excel's limitations on conditional format properties
//!
//! It is important to note that not all of Excel's cell format properties can
//! be modified with a conditional format.
//!
//! For example the view below of the Excel conditional format dialog shows the
//! limited number of font properties that can be set. The available properties
//! are highlighted with green.
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/conditional_format_limitations.png">
//!
//! Properties that **cannot** be modified in a conditional format are font
//! name, font size, superscript and subscript, diagonal borders, all alignment
//! properties and all protection properties.
//!
//! # Selecting a non-contiguous range
//!
//! In Excel it is possible to select several non-contiguous cells or ranges
//! like `"B3:D6 I3:K6 B9:D12 I9:K12"` and apply a conditional format to them.
//!
//! It is possible to achieve a similar effect with `rust_xlsxwriter` by using
//! repeated calls to
//! [`Worksheet::add_conditional_format()`](crate::Worksheet::add_conditional_format).
//! However this approach results in multiple identical conditional formats
//! applied to different cell ranges rather than one conditional format applied
//! to a multiple range selection.
//!
//! If this distinction is important to you then you can get the Excel multiple
//! selection effect using the `set_multi_range()` which is provided for all the
//! `ConditionalFormat` types. See the example below and note that the cells
//! outside the selected ranges do not have any conditional formatting.
//!
//!
//! Note, you can use an Excel range like
//! `"$B$3:$D$6,$I$3:$K$6,$B$9:$D$12,$I$9:$K$12"` or omit the `$` anchors and
//! replace the commas with spaces to have a clearer range like `"B3:D6 I3:K6
//! B9:D12 I9:K12"`. The documentation and examples use the latter format for
//! clarity but it you are copying and pasting from Excel you can use the first
//! format.
//!
//! ```
//! # // This code is available in examples/doc_conditional_format_multi_range.rs
//! #
//! # use rust_xlsxwriter::{
//! #     ConditionalFormatCell, ConditionalFormatCellCriteria, Format, Workbook, XlsxError,
//! # };
//! #
//! # fn main() -> Result<(), XlsxError> {
//! #     // Create a new Excel file object.
//! #     let mut workbook = Workbook::new();
//! #     let worksheet = workbook.add_worksheet();
//! #
//! #     // Add some sample data.
//! #     let data = [
//! #         [90, 80, 50, 10, 20, 90, 40, 90, 30, 40],
//! #         [20, 10, 90, 100, 30, 60, 70, 60, 50, 90],
//! #         [10, 50, 60, 50, 20, 50, 80, 30, 40, 60],
//! #         [10, 90, 20, 40, 10, 40, 50, 70, 90, 50],
//! #         [70, 100, 10, 90, 10, 10, 20, 100, 100, 40],
//! #         [20, 60, 10, 100, 30, 10, 20, 60, 100, 10],
//! #         [10, 60, 10, 80, 100, 80, 30, 30, 70, 40],
//! #         [30, 90, 60, 10, 10, 100, 40, 40, 30, 40],
//! #         [80, 90, 10, 20, 20, 50, 80, 20, 60, 90],
//! #         [60, 80, 30, 30, 10, 50, 80, 60, 50, 30],
//! #     ];
//! #     worksheet.write_row_matrix(2, 1, data)?;
//! #
//! #     // Set the column widths for clarity.
//! #     for col_num in 1..=10u16 {
//! #         worksheet.set_column_width(col_num, 6)?;
//! #     }
//! #
//! #     // Add a format. Light red fill with dark red text.
//! #     let format1 = Format::new()
//! #         .set_font_color("9C0006")
//! #         .set_background_color("FFC7CE");
//! #
//! #     // Add a format. Green fill with dark green text.
//! #     let format2 = Format::new()
//! #         .set_font_color("006100")
//! #         .set_background_color("C6EFCE");
//! #
//!     // Write a conditional format over a non-contiguous range.
//!     let conditional_format = ConditionalFormatCell::new()
//!         .set_criteria(ConditionalFormatCellCriteria::GreaterThanOrEqualTo)
//!         .set_value(50)
//!         .set_multi_range("B3:D6 I3:K6 B9:D12 I9:K12")
//!         .set_format(format1);
//!
//!     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
//!
//!     // Write another conditional format over the same range.
//!     let conditional_format = ConditionalFormatCell::new()
//!         .set_criteria(ConditionalFormatCellCriteria::LessThan)
//!         .set_value(50)
//!         .set_multi_range("B3:D6 I3:K6 B9:D12 I9:K12")
//!         .set_format(format2);
//!
//!     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
//!
//! #     // Save the file.
//! #     workbook.save("conditional_format.xlsx")?;
//! #
//! #     Ok(())
//! # }
//! ```
//!
//! Output file:
//!
//! <img
//! src="https://rustxlsxwriter.github.io/images/conditional_format_multi_range.png">
//!
#![warn(missing_docs)]

mod tests;

#[cfg(feature = "chrono")]
use chrono::{NaiveDate, NaiveDateTime, NaiveTime};

use crate::{xmlwriter::XMLWriter, ExcelDateTime, Format, Formula, XlsxError};
use std::{borrow::Cow, fmt};

// -----------------------------------------------------------------------
// ConditionalFormat trait
// -----------------------------------------------------------------------

/// Trait for generic conditional format types.
///
pub trait ConditionalFormat {
    /// Validate the conditional format.
    ///
    /// # Errors
    ///
    /// * [`XlsxError::ConditionalFormatError`] - A general error that is raised
    ///   when a conditional formatting parameter is incorrect or missing.
    ///
    fn validate(&self) -> Result<(), XlsxError>;

    /// Return the conditional format rule as an XML string.
    fn get_rule_string(&self, dxf_index: Option<u32>, priority: u32, anchor: &str) -> String;

    /// Get a mutable reference to the format object in the conditional format.
    fn format_as_mut(&mut self) -> Option<&mut Format>;

    /// Get the index of the format object in the conditional format.
    fn format_index(&self) -> Option<u32>;

    /// Get the multi-cell range for the conditional format, if present.
    fn multi_range(&self) -> String;

    /// Clone a reference into a concrete Box type.
    fn box_clone(&self) -> Box<dyn ConditionalFormat + Send>;
}

macro_rules! generate_conditional_format_impls {
    ($($t:ty)*) => ($(
        impl ConditionalFormat for $t {
            fn validate(&self) -> Result<(), XlsxError> {
                self.validate()
            }

            fn get_rule_string(&self, dxf_index: Option<u32>, priority: u32, anchor: &str) -> String {
                self.get_rule_string(dxf_index, priority, anchor)
            }

            fn format_as_mut(&mut self) -> Option<&mut Format> {
                self.format_as_mut()
            }

            fn format_index(&self) -> Option<u32> {
                self.format_index()
            }

            fn multi_range(&self) -> String {
                self.multi_range()
            }

            fn box_clone(&self) -> Box<dyn ConditionalFormat + Send> {
                Box::new(self.clone())
            }
        }
    )*)
}
generate_conditional_format_impls!(
    ConditionalFormatAverage
    ConditionalFormatCell
    ConditionalFormatDuplicate
    ConditionalFormatText
    ConditionalFormatTop
);

// -----------------------------------------------------------------------
// ConditionalFormatCell
// -----------------------------------------------------------------------

/// The `ConditionalFormatCell` struct represents a Cell conditional format.
///
/// `ConditionalFormatCell` is used to represent a Cell style conditional format
/// in Excel. Cell conditional formats use simple equalities such as "equal to"
/// or "greater than" or "between".
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/conditional_format_cell_intro.png">
///
/// For more information see [Working with Conditional Formats](crate::conditional_format).
///
/// # Examples
///
/// Example of adding a cell type conditional formatting to a worksheet. Cells
/// with values >= 50 are in light red. Values < 50 are in light green.
///
/// ```
/// # // This code is available in examples/doc_conditional_format_cell1.rs
/// #
/// # use rust_xlsxwriter::{
/// #     ConditionalFormatCell, ConditionalFormatCellCriteria, Format, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some sample data.
/// #     let data = [
/// #         [90, 80, 50, 10, 20, 90, 40, 90, 30, 40],
/// #         [20, 10, 90, 100, 30, 60, 70, 60, 50, 90],
/// #         [10, 50, 60, 50, 20, 50, 80, 30, 40, 60],
/// #         [10, 90, 20, 40, 10, 40, 50, 70, 90, 50],
/// #         [70, 100, 10, 90, 10, 10, 20, 100, 100, 40],
/// #         [20, 60, 10, 100, 30, 10, 20, 60, 100, 10],
/// #         [10, 60, 10, 80, 100, 80, 30, 30, 70, 40],
/// #         [30, 90, 60, 10, 10, 100, 40, 40, 30, 40],
/// #         [80, 90, 10, 20, 20, 50, 80, 20, 60, 90],
/// #         [60, 80, 30, 30, 10, 50, 80, 60, 50, 30],
/// #     ];
/// #     worksheet.write_row_matrix(2, 1, data)?;
/// #
/// #     // Set the column widths for clarity.
/// #     for col_num in 1..=10u16 {
/// #         worksheet.set_column_width(col_num, 6)?;
/// #     }
/// #
/// #     // Add a format. Light red fill with dark red text.
/// #     let format1 = Format::new()
/// #         .set_font_color("9C0006")
/// #         .set_background_color("FFC7CE");
/// #
/// #     // Add a format. Green fill with dark green text.
/// #     let format2 = Format::new()
/// #         .set_font_color("006100")
/// #         .set_background_color("C6EFCE");
/// #
///     // Write a conditional format over a range.
///     let conditional_format = ConditionalFormatCell::new()
///         .set_criteria(ConditionalFormatCellCriteria::GreaterThanOrEqualTo)
///         .set_value(50)
///         .set_format(format1);
///
///     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
///
///     // Write another conditional format over the same range.
///     let conditional_format = ConditionalFormatCell::new()
///         .set_criteria(ConditionalFormatCellCriteria::LessThan)
///         .set_value(50)
///         .set_format(format2);
///
///     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
///
/// #     // Save the file.
/// #     workbook.save("conditional_format.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/conditional_format_cell1.png">
///
///
/// Example of adding a cell type conditional formatting to a worksheet. Values
/// between 30 and 70 are highlighted in light red. Values outside that range
/// are in light green.
///
/// ```
/// # // This code is available in examples/doc_conditional_format_cell2.rs
/// #
/// # use rust_xlsxwriter::{
/// #     ConditionalFormatCell, ConditionalFormatCellCriteria, Format, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some sample data.
/// #     let data = [
/// #         [90, 80, 50, 10, 20, 90, 40, 90, 30, 40],
/// #         [20, 10, 90, 100, 30, 60, 70, 60, 50, 90],
/// #         [10, 50, 60, 50, 20, 50, 80, 30, 40, 60],
/// #         [10, 90, 20, 40, 10, 40, 50, 70, 90, 50],
/// #         [70, 100, 10, 90, 10, 10, 20, 100, 100, 40],
/// #         [20, 60, 10, 100, 30, 10, 20, 60, 100, 10],
/// #         [10, 60, 10, 80, 100, 80, 30, 30, 70, 40],
/// #         [30, 90, 60, 10, 10, 100, 40, 40, 30, 40],
/// #         [80, 90, 10, 20, 20, 50, 80, 20, 60, 90],
/// #         [60, 80, 30, 30, 10, 50, 80, 60, 50, 30],
/// #     ];
/// #     worksheet.write_row_matrix(2, 1, data)?;
/// #
/// #     // Set the column widths for clarity.
/// #     for col_num in 1..=10u16 {
/// #         worksheet.set_column_width(col_num, 6)?;
/// #     }
/// #
/// #     // Add a format. Light red fill with dark red text.
/// #     let format1 = Format::new()
/// #         .set_font_color("9C0006")
/// #         .set_background_color("FFC7CE");
/// #
/// #     // Add a format. Green fill with dark green text.
/// #     let format2 = Format::new()
/// #         .set_font_color("006100")
/// #         .set_background_color("C6EFCE");
/// #
///     // Write a conditional format over a range.
///     let conditional_format = ConditionalFormatCell::new()
///         .set_criteria(ConditionalFormatCellCriteria::Between)
///         .set_minimum(30)
///         .set_maximum(70)
///         .set_format(format1);
///
///     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
///
///     // Write another conditional format over the same range.
///     let conditional_format = ConditionalFormatCell::new()
///         .set_criteria(ConditionalFormatCellCriteria::NotBetween)
///         .set_minimum(30)
///         .set_maximum(70)
///         .set_format(format2);
///
///     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
///
/// #     // Save the file.
/// #     workbook.save("conditional_format.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/conditional_format_cell2.png">
///
#[derive(Clone)]
pub struct ConditionalFormatCell {
    minimum: ConditionalFormatValue,
    maximum: ConditionalFormatValue,
    criteria: ConditionalFormatCellCriteria,
    multi_range: String,
    stop_if_true: bool,
    pub(crate) format: Option<Format>,
}

/// **Section 1**: The following methods are specific to `ConditionalFormatCell`.
impl ConditionalFormatCell {
    /// Create a new Cell conditional format struct.
    #[allow(clippy::new_without_default)]
    pub fn new() -> ConditionalFormatCell {
        ConditionalFormatCell {
            minimum: ConditionalFormatValue::new_from_string(""),
            maximum: ConditionalFormatValue::new_from_string(""),
            criteria: ConditionalFormatCellCriteria::None,
            multi_range: String::new(),
            stop_if_true: false,
            format: None,
        }
    }

    /// Set the value of the Cell conditional format rule.
    ///
    /// # Parameters
    ///
    /// * `value` - Any type that can convert into a [`ConditionalFormatValue`]
    ///   which is effectively all types supported by Excel.
    ///
    /// # Examples
    ///
    /// Example of adding a cell type conditional formatting to a worksheet. Cells
    /// with values >= 50 are in light green.
    ///
    /// ```
    /// # // This code is available in examples/doc_conditional_format_cell_set_value.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     ConditionalFormatCell, ConditionalFormatCellCriteria, Format, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some sample data.
    /// #     let data = [10, 80, 50, 10, 20, 60, 40, 70, 30, 40];
    /// #
    /// #     worksheet.write_column(0, 0, data)?;
    /// #
    /// #     // Add a format. Green fill with dark green text.
    /// #     let format = Format::new()
    /// #         .set_font_color("006100")
    /// #         .set_background_color("C6EFCE");
    /// #
    ///     // Write a conditional format over a range.
    ///     let conditional_format = ConditionalFormatCell::new()
    ///         .set_criteria(ConditionalFormatCellCriteria::GreaterThanOrEqualTo)
    ///         .set_value(50)
    ///         .set_format(format);
    ///
    ///     worksheet.add_conditional_format(0, 0, 9, 0, &conditional_format)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("conditional_format.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/conditional_format_cell_set_value.png">
    ///
    pub fn set_value(self, value: impl Into<ConditionalFormatValue>) -> ConditionalFormatCell {
        self.set_minimum(value)
    }

    /// Set the minimum value of the Cell "between" and "not between"
    /// conditional format rules.
    ///
    /// # Parameters
    ///
    /// * `value` - Any type that can convert into a [`ConditionalFormatValue`]
    ///   which is effectively all types supported by Excel.
    ///
    /// # Examples
    ///
    /// Example of adding a cell type conditional formatting to a worksheet.
    /// Values between 40 and 60 are highlighted in light green.
    ///
    /// ```
    /// # // This code is available in examples/doc_conditional_format_cell_set_minimum.rs
    /// #
    /// # use rust_xlsxwriter::{
    /// #     ConditionalFormatCell, ConditionalFormatCellCriteria, Format, Workbook, XlsxError,
    /// # };
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     // Add some sample data.
    /// #     let data = [10, 80, 50, 10, 20, 60, 40, 70, 30, 40];
    /// #
    /// #     worksheet.write_column(0, 0, data)?;
    /// #
    /// #     // Add a format. Green fill with dark green text.
    /// #     let format = Format::new()
    /// #         .set_font_color("006100")
    /// #         .set_background_color("C6EFCE");
    /// #
    ///     // Write a conditional format over a range.
    ///     let conditional_format = ConditionalFormatCell::new()
    ///         .set_criteria(ConditionalFormatCellCriteria::Between)
    ///         .set_minimum(40)
    ///         .set_maximum(60)
    ///         .set_format(format);
    ///
    ///     worksheet.add_conditional_format(0, 0, 9, 0, &conditional_format)?;
    ///
    /// #     // Save the file.
    /// #     workbook.save("conditional_format.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/conditional_format_cell_set_minimum.png">
    ///
    pub fn set_minimum(
        mut self,
        value: impl Into<ConditionalFormatValue>,
    ) -> ConditionalFormatCell {
        self.minimum = value.into();
        self.minimum.quote_string();
        self
    }

    /// Set the maximum value of the Cell "between" and "not between"
    /// conditional format rules.
    ///
    /// Set the example above.
    ///
    /// # Parameters
    ///
    /// * `value` - Any type that can convert into a [`ConditionalFormatValue`]
    ///   which is effectively all types supported by Excel.
    ///
    pub fn set_maximum(
        mut self,
        value: impl Into<ConditionalFormatValue>,
    ) -> ConditionalFormatCell {
        self.maximum = value.into();
        self.maximum.quote_string();
        self
    }

    /// Set the criteria for the conditional format rule such as `=`, `!=`, `>`,
    /// `<`, `>=`, `<=`, `between` or `not between`.
    ///
    /// # Parameters
    ///
    /// * `criteria` - A [`ConditionalFormatCellCriteria`] enum value.
    ///
    pub fn set_criteria(
        mut self,
        criteria: ConditionalFormatCellCriteria,
    ) -> ConditionalFormatCell {
        self.criteria = criteria;
        self
    }

    // Validate the conditional format.
    pub(crate) fn validate(&self) -> Result<(), XlsxError> {
        let error_message = match self.criteria {
            ConditionalFormatCellCriteria::None => "'criteria' must be set".to_string(),
            ConditionalFormatCellCriteria::Between | ConditionalFormatCellCriteria::NotBetween => {
                if self.minimum.value.is_empty() {
                    "'minimum' value must be set".to_string()
                } else if self.maximum.value.is_empty() {
                    "'maximum' value must be set".to_string()
                } else {
                    String::new()
                }
            }
            _ => {
                if self.minimum.value.is_empty() {
                    "'value' must be set".to_string()
                } else {
                    String::new()
                }
            }
        };

        if !error_message.is_empty() {
            return Err(XlsxError::ConditionalFormatError(error_message));
        }

        Ok(())
    }

    //  Return the conditional format rule as an XML string.
    pub(crate) fn get_rule_string(
        &self,
        dxf_index: Option<u32>,
        priority: u32,
        _anchor: &str,
    ) -> String {
        let mut writer = XMLWriter::new();
        let mut attributes = vec![("type", "cellIs".to_string())];

        // Set the format index if present.
        if let Some(dxf_index) = dxf_index {
            attributes.push(("dxfId", dxf_index.to_string()));
        }

        // Set the rule priority order.
        attributes.push(("priority", priority.to_string()));

        // Set the "Stop if True" property.
        if self.stop_if_true {
            attributes.push(("stopIfTrue", "1".to_string()));
        }

        attributes.push(("operator", self.criteria.to_string()));

        // Write the rule.
        writer.xml_start_tag("cfRule", &attributes);
        writer.xml_data_element_only("formula", &self.minimum.value);

        if self.criteria == ConditionalFormatCellCriteria::Between
            || self.criteria == ConditionalFormatCellCriteria::NotBetween
        {
            writer.xml_data_element_only("formula", &self.maximum.value);
        }

        writer.xml_end_tag("cfRule");

        writer.read_to_string()
    }
}

// -----------------------------------------------------------------------
// ConditionalFormatDuplicate
// -----------------------------------------------------------------------

/// The `ConditionalFormatDuplicate` struct represents a Duplicate/Unique
/// conditional format.
///
/// `ConditionalFormatDuplicate` is used to represent a Duplicate or Unique
/// style conditional format in Excel. Duplicate conditional formats show
/// duplicated values in a range while Unique conditional formats show unique
/// values.
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/conditional_format_duplicate_intro.png">
///
/// For more information see [Working with Conditional Formats](crate::conditional_format).
///
/// # Examples
///
/// Example of how to add a duplicate/unique conditional formatting to a
/// worksheet. Duplicate values are in light red. Unique values are in light
/// green. Note, that we invert the Duplicate rule to get Unique values.
///
/// ```
/// # // This code is available in examples/doc_conditional_format_duplicate.rs
/// #
/// # use rust_xlsxwriter::{ConditionalFormatDuplicate, Format, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some sample data.
/// #     let data = [
/// #         [34, 72, 38, 30, 75, 48, 75, 66, 84, 86],
/// #         [6, 24, 1, 84, 54, 62, 60, 3, 26, 59],
/// #         [28, 79, 97, 13, 85, 93, 93, 22, 5, 14],
/// #         [27, 71, 40, 17, 18, 79, 90, 93, 29, 47],
/// #         [88, 25, 33, 23, 67, 1, 59, 79, 47, 36],
/// #         [24, 100, 20, 88, 29, 33, 38, 54, 54, 88],
/// #         [6, 57, 88, 28, 10, 26, 37, 7, 41, 48],
/// #         [52, 78, 1, 96, 26, 45, 47, 33, 96, 36],
/// #         [60, 54, 81, 66, 81, 90, 80, 93, 12, 55],
/// #         [70, 5, 46, 14, 71, 19, 66, 36, 41, 21],
/// #     ];
/// #     worksheet.write_row_matrix(2, 1, data)?;
/// #
/// #     // Set the column widths for clarity.
/// #     for col_num in 1..=10u16 {
/// #         worksheet.set_column_width(col_num, 6)?;
/// #     }
/// #
/// #     // Add a format. Light red fill with dark red text.
/// #     let format1 = Format::new()
/// #         .set_font_color("9C0006")
/// #         .set_background_color("FFC7CE");
/// #
/// #     // Add a format. Green fill with dark green text.
/// #     let format2 = Format::new()
/// #         .set_font_color("006100")
/// #         .set_background_color("C6EFCE");
/// #
///     // Write a conditional format over a range.
///     let conditional_format = ConditionalFormatDuplicate::new().set_format(format1);
///
///     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
///
///     // Invert the duplicate conditional format to show unique values.
///     let conditional_format = ConditionalFormatDuplicate::new()
///         .invert()
///         .set_format(format2);
///
///     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
///
/// #     // Save the file.
/// #     workbook.save("conditional_format.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/conditional_format_duplicate.png">
///
#[derive(Clone)]
pub struct ConditionalFormatDuplicate {
    is_unique: bool,
    multi_range: String,
    stop_if_true: bool,
    pub(crate) format: Option<Format>,
}

/// **Section 1**: The following methods are specific to `ConditionalFormatDuplicate`.
impl ConditionalFormatDuplicate {
    /// Create a new Duplicate conditional format struct.
    #[allow(clippy::new_without_default)]
    pub fn new() -> ConditionalFormatDuplicate {
        ConditionalFormatDuplicate {
            is_unique: false,
            multi_range: String::new(),
            stop_if_true: false,
            format: None,
        }
    }

    /// Invert the functionality of the conditional format to get unique values
    /// instead of duplicate values.
    ///
    /// See the example above.
    ///
    pub fn invert(mut self) -> ConditionalFormatDuplicate {
        self.is_unique = true;
        self
    }

    // Validate the conditional format.
    #[allow(clippy::unnecessary_wraps)]
    #[allow(clippy::unused_self)]
    pub(crate) fn validate(&self) -> Result<(), XlsxError> {
        Ok(())
    }

    //  Return the conditional format rule as an XML string.
    pub(crate) fn get_rule_string(
        &self,
        dxf_index: Option<u32>,
        priority: u32,
        _anchor: &str,
    ) -> String {
        let mut writer = XMLWriter::new();
        let mut attributes = vec![];

        if self.is_unique {
            attributes.push(("type", "uniqueValues".to_string()));
        } else {
            attributes.push(("type", "duplicateValues".to_string()));
        }

        // Set the format index if present.
        if let Some(dxf_index) = dxf_index {
            attributes.push(("dxfId", dxf_index.to_string()));
        }

        // Set the rule priority order.
        attributes.push(("priority", priority.to_string()));

        // Set the "Stop if True" property.
        if self.stop_if_true {
            attributes.push(("stopIfTrue", "1".to_string()));
        }

        // Write the rule.
        writer.xml_empty_tag("cfRule", &attributes);

        writer.read_to_string()
    }
}

// -----------------------------------------------------------------------
// ConditionalFormatAverage
// -----------------------------------------------------------------------

/// The `ConditionalFormatAverage` struct represents an Average/Standard
/// Deviation style conditional format.
///
/// `ConditionalFormatAverage` is used to represent a Average or Standard Deviation style
/// conditional format in Excel.
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/conditional_format_average_intro.png">
///
/// For more information see [Working with Conditional
/// Formats](crate::conditional_format).
///
/// # Examples
///
/// Example of how to add Average conditional formatting to a worksheet. Above
/// average values are in light red. Below average values are in light green.
///
/// ```
/// # // This code is available in examples/doc_conditional_format_average.rs
/// #
/// # use rust_xlsxwriter::{
/// #     ConditionalFormatAverage, ConditionalFormatAverageCriteria, Format, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some sample data.
/// #     let data = [
/// #         [34, 72, 38, 30, 75, 48, 75, 66, 84, 86],
/// #         [6, 24, 1, 84, 54, 62, 60, 3, 26, 59],
/// #         [28, 79, 97, 13, 85, 93, 93, 22, 5, 14],
/// #         [27, 71, 40, 17, 18, 79, 90, 93, 29, 47],
/// #         [88, 25, 33, 23, 67, 1, 59, 79, 47, 36],
/// #         [24, 100, 20, 88, 29, 33, 38, 54, 54, 88],
/// #         [6, 57, 88, 28, 10, 26, 37, 7, 41, 48],
/// #         [52, 78, 1, 96, 26, 45, 47, 33, 96, 36],
/// #         [60, 54, 81, 66, 81, 90, 80, 93, 12, 55],
/// #         [70, 5, 46, 14, 71, 19, 66, 36, 41, 21],
/// #     ];
/// #     worksheet.write_row_matrix(2, 1, data)?;
/// #
/// #     // Set the column widths for clarity.
/// #     for col_num in 1..=10u16 {
/// #         worksheet.set_column_width(col_num, 6)?;
/// #     }
/// #
/// #     // Add a format. Light red fill with dark red text.
/// #     let format1 = Format::new()
/// #         .set_font_color("9C0006")
/// #         .set_background_color("FFC7CE");
/// #
/// #     // Add a format. Green fill with dark green text.
/// #     let format2 = Format::new()
/// #         .set_font_color("006100")
/// #         .set_background_color("C6EFCE");
/// #
///     // Write a conditional format. The default criteria is Above Average.
///     let conditional_format = ConditionalFormatAverage::new().set_format(format1);
///
///     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
///
///     // Write another conditional format over the same range.
///     let conditional_format = ConditionalFormatAverage::new()
///         .set_criteria(ConditionalFormatAverageCriteria::BelowAverage)
///         .set_format(format2);
///
///     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
///
/// #     // Save the file.
/// #     workbook.save("conditional_format.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/conditional_format_average.png">
///
#[derive(Clone)]
pub struct ConditionalFormatAverage {
    criteria: ConditionalFormatAverageCriteria,
    multi_range: String,
    stop_if_true: bool,
    pub(crate) format: Option<Format>,
}

/// **Section 1**: The following methods are specific to `ConditionalFormatAverage`.
impl ConditionalFormatAverage {
    /// Create a new Average conditional format struct.
    #[allow(clippy::new_without_default)]
    pub fn new() -> ConditionalFormatAverage {
        ConditionalFormatAverage {
            criteria: ConditionalFormatAverageCriteria::AboveAverage,
            multi_range: String::new(),
            stop_if_true: false,
            format: None,
        }
    }

    /// Set the criteria for the Average conditional format rule such as above
    /// or below average or standard deviation ranges.
    ///
    /// # Parameters
    ///
    /// * `criteria` - A [`ConditionalFormatAverageCriteria`] enum value.
    ///
    pub fn set_criteria(
        mut self,
        criteria: ConditionalFormatAverageCriteria,
    ) -> ConditionalFormatAverage {
        self.criteria = criteria;
        self
    }

    // Validate the conditional format.
    #[allow(clippy::unnecessary_wraps)]
    #[allow(clippy::unused_self)]
    pub(crate) fn validate(&self) -> Result<(), XlsxError> {
        Ok(())
    }

    //  Return the conditional format rule as an XML string.
    pub(crate) fn get_rule_string(
        &self,
        dxf_index: Option<u32>,
        priority: u32,
        _anchor: &str,
    ) -> String {
        let mut writer = XMLWriter::new();
        let mut attributes = vec![("type", "aboveAverage".to_string())];

        // Set the format index if present.
        if let Some(dxf_index) = dxf_index {
            attributes.push(("dxfId", dxf_index.to_string()));
        }

        // Set the rule priority order.
        attributes.push(("priority", priority.to_string()));

        // Set the "Stop if True" property.
        if self.stop_if_true {
            attributes.push(("stopIfTrue", "1".to_string()));
        }

        // Set the Average specific attributes.
        match self.criteria {
            ConditionalFormatAverageCriteria::AboveAverage => {
                // There are no additional attributes for above average.
            }

            ConditionalFormatAverageCriteria::BelowAverage => {
                attributes.push(("aboveAverage", "0".to_string()));
            }

            ConditionalFormatAverageCriteria::EqualOrAboveAverage => {
                attributes.push(("equalAverage", "1".to_string()));
            }

            ConditionalFormatAverageCriteria::EqualOrBelowAverage => {
                attributes.push(("aboveAverage", "0".to_string()));
                attributes.push(("equalAverage", "1".to_string()));
            }

            ConditionalFormatAverageCriteria::OneStandardDeviationAbove => {
                attributes.push(("stdDev", "1".to_string()));
            }

            ConditionalFormatAverageCriteria::OneStandardDeviationBelow => {
                attributes.push(("aboveAverage", "0".to_string()));
                attributes.push(("stdDev", "1".to_string()));
            }

            ConditionalFormatAverageCriteria::TwoStandardDeviationsAbove => {
                attributes.push(("stdDev", "2".to_string()));
            }

            ConditionalFormatAverageCriteria::TwoStandardDeviationsBelow => {
                attributes.push(("aboveAverage", "0".to_string()));

                attributes.push(("stdDev", "2".to_string()));
            }

            ConditionalFormatAverageCriteria::ThreeStandardDeviationsAbove => {
                attributes.push(("stdDev", "3".to_string()));
            }

            ConditionalFormatAverageCriteria::ThreeStandardDeviationsBelow => {
                attributes.push(("aboveAverage", "0".to_string()));

                attributes.push(("stdDev", "3".to_string()));
            }
        }

        // Write the rule.
        writer.xml_empty_tag("cfRule", &attributes);

        writer.read_to_string()
    }
}

// -----------------------------------------------------------------------
// ConditionalFormatTop
// -----------------------------------------------------------------------

/// The `ConditionalFormatTop` struct represents a Top/Bottom style conditional
/// format.
///
/// `ConditionalFormatTop` is used to represent a Top or Bottom style
/// conditional format in Excel. Top conditional formats show the top X values
/// in a range. The value of the conditional can be a rank, i.e., Top X, or a
/// percentage, i.e., Top X%.
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/conditional_format_top_intro.png">
///
/// For more information see [Working with Conditional
/// Formats](crate::conditional_format).
///
/// # Examples
///
/// Example of how to add Top and Bottom conditional formatting to a worksheet.
/// Top 10 values are in light red. Bottom 10 values are in light green.
///
/// ```
/// # // This code is available in examples/doc_conditional_format_top.rs
/// #
/// # use rust_xlsxwriter::{ConditionalFormatTop, Format, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some sample data.
/// #     let data = [
/// #         [34, 72, 38, 30, 75, 48, 75, 66, 84, 86],
/// #         [6, 24, 1, 84, 54, 62, 60, 3, 26, 59],
/// #         [28, 79, 97, 13, 85, 93, 93, 22, 5, 14],
/// #         [27, 71, 40, 17, 18, 79, 90, 93, 29, 47],
/// #         [88, 25, 33, 23, 67, 1, 59, 79, 47, 36],
/// #         [24, 100, 20, 88, 29, 33, 38, 54, 54, 88],
/// #         [6, 57, 88, 28, 10, 26, 37, 7, 41, 48],
/// #         [52, 78, 1, 96, 26, 45, 47, 33, 96, 36],
/// #         [60, 54, 81, 66, 81, 90, 80, 93, 12, 55],
/// #         [70, 5, 46, 14, 71, 19, 66, 36, 41, 21],
/// #     ];
/// #     worksheet.write_row_matrix(2, 1, data)?;
/// #
/// #     // Set the column widths for clarity.
/// #     for col_num in 1..=10u16 {
/// #         worksheet.set_column_width(col_num, 6)?;
/// #     }
/// #
/// #     // Add a format. Light red fill with dark red text.
/// #     let format1 = Format::new()
/// #         .set_font_color("9C0006")
/// #         .set_background_color("FFC7CE");
/// #
/// #     // Add a format. Green fill with dark green text.
/// #     let format2 = Format::new()
/// #         .set_font_color("006100")
/// #         .set_background_color("C6EFCE");
/// #
///     // Write a conditional format over a range.
///     let conditional_format = ConditionalFormatTop::new()
///         .set_value(10)
///         .set_format(format1);
///
///     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
///
///     // Invert the Top conditional format to show Bottom values.
///     let conditional_format = ConditionalFormatTop::new()
///         .invert()
///         .set_value(10)
///         .set_format(format2);
///     worksheet.add_conditional_format(2, 1, 11, 10, &conditional_format)?;
///
/// #     // Save the file.
/// #     workbook.save("conditional_format.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/conditional_format_top.png">
///
#[derive(Clone)]
pub struct ConditionalFormatTop {
    value: u16,
    is_bottom: bool,
    is_percent: bool,
    multi_range: String,
    stop_if_true: bool,
    pub(crate) format: Option<Format>,
}

/// **Section 1**: The following methods are specific to `ConditionalFormatTop`.
impl ConditionalFormatTop {
    /// Create a new Top conditional format struct.
    #[allow(clippy::new_without_default)]
    pub fn new() -> ConditionalFormatTop {
        ConditionalFormatTop {
            value: 10,
            is_bottom: false,
            is_percent: false,
            multi_range: String::new(),
            stop_if_true: false,
            format: None,
        }
    }

    /// Set the top/bottom rank of the conditional format.
    ///
    /// See the example above.
    ///
    pub fn set_value(mut self, value: u16) -> ConditionalFormatTop {
        self.value = value;
        self
    }

    /// Invert the functionality of the conditional format to the get Bottom values
    /// instead of the Top values.
    ///
    /// See the example above.
    ///
    pub fn invert(mut self) -> ConditionalFormatTop {
        self.is_bottom = true;
        self
    }

    /// Show the top/bottom percentage instead of rank.
    ///
    /// See the example above.
    ///
    pub fn set_percent(mut self) -> ConditionalFormatTop {
        self.is_percent = true;
        self
    }

    // Validate the conditional format.
    pub(crate) fn validate(&self) -> Result<(), XlsxError> {
        if !(1..=1000).contains(&self.value) {
            return Err(XlsxError::ConditionalFormatError(
                "value must be in the Excel range 1..1000".to_string(),
            ));
        }

        Ok(())
    }

    //  Return the conditional format rule as an XML string.
    pub(crate) fn get_rule_string(
        &self,
        dxf_index: Option<u32>,
        priority: u32,
        _anchor: &str,
    ) -> String {
        let mut writer = XMLWriter::new();
        let mut attributes = vec![("type", "top10".to_string())];

        // Set the format index if present.
        if let Some(dxf_index) = dxf_index {
            attributes.push(("dxfId", dxf_index.to_string()));
        }

        // Set the rule priority order.
        attributes.push(("priority", priority.to_string()));

        // Set the "Stop if True" property.
        if self.stop_if_true {
            attributes.push(("stopIfTrue", "1".to_string()));
        }

        if self.is_percent {
            attributes.push(("percent", "1".to_string()));
        }

        if self.is_bottom {
            attributes.push(("bottom", "1".to_string()));
        }

        attributes.push(("rank", self.value.to_string()));

        // Write the rule.
        writer.xml_empty_tag("cfRule", &attributes);

        writer.read_to_string()
    }
}

// -----------------------------------------------------------------------
// ConditionalFormatText
// -----------------------------------------------------------------------

/// The `ConditionalFormatText` struct represents a Text conditional format.
///
/// `ConditionalFormatText` is used to represent a Text style conditional format
/// in Excel. Text conditional formats use simple equalities such as "equal to"
/// or "greater than" or "between".
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/conditional_format_text_intro.png">
///
/// For more information see [Working with Conditional Formats](crate::conditional_format).
///
/// # Examples
///
/// Example of adding a text type conditional formatting to a worksheet.
///
/// ```
/// # // This code is available in examples/doc_conditional_format_text.rs
/// #
/// # use rust_xlsxwriter::{
/// #     ConditionalFormatText, ConditionalFormatTextCriteria, Format, Workbook, XlsxError,
/// # };
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     // Add some sample data.
/// #     let data = [
/// #         "apocrustic",
/// #         "burstwort",
/// #         "cloudburst",
/// #         "crustification",
/// #         "distrustfulness",
/// #         "laurustine",
/// #         "outburst",
/// #         "rusticism",
/// #         "thunderburst",
/// #         "trustee",
/// #         "trustworthiness",
/// #         "unburstableness",
/// #         "unfrustratable",
/// #     ];
/// #     worksheet.write_column(0, 0, data)?;
/// #     worksheet.write_column(0, 2, data)?;
/// #
/// #     // Set the column widths for clarity.
/// #     worksheet.set_column_width(0, 20)?;
/// #     worksheet.set_column_width(2, 20)?;
/// #
/// #     // Add a format. Green fill with dark green text.
/// #     let format1 = Format::new()
/// #         .set_font_color("006100")
/// #         .set_background_color("C6EFCE");
/// #
/// #     // Add a format. Light red fill with dark red text.
/// #     let format2 = Format::new()
/// #         .set_font_color("9C0006")
/// #         .set_background_color("FFC7CE");
/// #
///     // Write a text "containing" conditional format over a range.
///     let conditional_format = ConditionalFormatText::new()
///         .set_criteria(ConditionalFormatTextCriteria::Contains)
///         .set_value("rust")
///         .set_format(&format1);
///
///     worksheet.add_conditional_format(0, 0, 12, 0, &conditional_format)?;
///
///     // Write a text "not containing" conditional format over the same range.
///     let conditional_format = ConditionalFormatText::new()
///         .set_criteria(ConditionalFormatTextCriteria::DoesNotContain)
///         .set_value("rust")
///         .set_format(&format2);
///
///     worksheet.add_conditional_format(0, 0, 12, 0, &conditional_format)?;
///
///     // Write a text "begins with" conditional format over a range.
///     let conditional_format = ConditionalFormatText::new()
///         .set_criteria(ConditionalFormatTextCriteria::BeginsWith)
///         .set_value("t")
///         .set_format(&format1);
///
///     worksheet.add_conditional_format(0, 2, 12, 2, &conditional_format)?;
///
///     // Write a text "ends with" conditional format over the same range.
///     let conditional_format = ConditionalFormatText::new()
///         .set_criteria(ConditionalFormatTextCriteria::EndsWith)
///         .set_value("t")
///         .set_format(&format2);
///
///     worksheet.add_conditional_format(0, 2, 12, 2, &conditional_format)?;
/// #
/// #     // Save the file.
/// #     workbook.save("conditional_format.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// This creates conditional format rules like this:
///
/// <img src="https://rustxlsxwriter.github.io/images/conditional_format_text_rules.png">
///
/// And the following output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/conditional_format_text.png">
///
#[derive(Clone)]
pub struct ConditionalFormatText {
    value: String,
    criteria: ConditionalFormatTextCriteria,
    multi_range: String,
    stop_if_true: bool,
    pub(crate) format: Option<Format>,
}

/// **Section 1**: The following methods are specific to `ConditionalFormatText`.
impl ConditionalFormatText {
    /// Create a new Text conditional format struct.
    #[allow(clippy::new_without_default)]
    pub fn new() -> ConditionalFormatText {
        ConditionalFormatText {
            value: String::new(),
            criteria: ConditionalFormatTextCriteria::Contains,
            multi_range: String::new(),
            stop_if_true: false,
            format: None,
        }
    }

    /// Set the value of the Text conditional format rule.
    ///
    /// # Parameters
    ///
    /// * `value` - A string like value.
    ///
    ///   Newer versions of Excel support support using a cell reference for the
    ///   value but that isn't currently supported by `rust_xlsxwriter`.
    ///
    pub fn set_value(mut self, value: impl Into<String>) -> ConditionalFormatText {
        self.value = value.into();
        self
    }

    /// Set the criteria for the Text conditional format rule such "contains" or
    /// "starts with".
    ///
    /// # Parameters
    ///
    /// * `criteria` - A [`ConditionalFormatTextCriteria`] enum value.
    ///
    pub fn set_criteria(
        mut self,
        criteria: ConditionalFormatTextCriteria,
    ) -> ConditionalFormatText {
        self.criteria = criteria;
        self
    }

    // Validate the conditional format.
    pub(crate) fn validate(&self) -> Result<(), XlsxError> {
        if self.value.is_empty() {
            return Err(XlsxError::ConditionalFormatError(
                "Text conditional format string cannot be empty".to_string(),
            ));
        }

        Ok(())
    }

    //  Return the conditional format rule as an XML string.
    pub(crate) fn get_rule_string(
        &self,
        dxf_index: Option<u32>,
        priority: u32,
        anchor: &str,
    ) -> String {
        let mut writer = XMLWriter::new();
        let mut attributes = vec![];
        let text = self.value.clone();
        let operator;
        let formula;

        // Set the rule attributes based on the criteria.
        match self.criteria {
            ConditionalFormatTextCriteria::Contains => {
                attributes.push(("type", "containsText".to_string()));
                operator = "containsText".to_string();
                formula = format!(r#"NOT(ISERROR(SEARCH("{text}",{anchor})))"#);
            }
            ConditionalFormatTextCriteria::DoesNotContain => {
                attributes.push(("type", "notContainsText".to_string()));
                operator = "notContains".to_string();
                formula = format!(r#"ISERROR(SEARCH("{text}",{anchor}))"#);
            }
            ConditionalFormatTextCriteria::BeginsWith => {
                attributes.push(("type", "beginsWith".to_string()));
                operator = "beginsWith".to_string();
                formula = format!(r#"LEFT({anchor},1)="{text}""#);
            }
            ConditionalFormatTextCriteria::EndsWith => {
                attributes.push(("type", "endsWith".to_string()));
                operator = "endsWith".to_string();
                formula = format!(r#"RIGHT({anchor},1)="{text}""#);
            }
        }

        // Set the format index if present.
        if let Some(dxf_index) = dxf_index {
            attributes.push(("dxfId", dxf_index.to_string()));
        }

        // Set the rule priority order.
        attributes.push(("priority", priority.to_string()));

        // Set the "Stop if True" property.
        if self.stop_if_true {
            attributes.push(("stopIfTrue", "1".to_string()));
        }

        // Add the attributes.
        attributes.push(("operator", operator));
        attributes.push(("text", text));

        writer.xml_start_tag("cfRule", &attributes);
        writer.xml_data_element_only("formula", &formula);
        writer.xml_end_tag("cfRule");

        writer.read_to_string()
    }
}

// -----------------------------------------------------------------------
// ConditionalFormatValue
// -----------------------------------------------------------------------

/// The `ConditionalFormatValue` struct represents a Cell conditional format
/// value types.
///
/// TODO - Explain `ConditionalFormatValue`
///
#[derive(Clone)]
pub struct ConditionalFormatValue {
    value: String,
    pub(crate) is_string: bool,
}

impl ConditionalFormatValue {
    pub(crate) fn new_from_string(value: impl Into<String>) -> ConditionalFormatValue {
        ConditionalFormatValue {
            value: value.into(),
            is_string: false,
        }
    }

    // Helper method to account for the fact that Excel requires that strings in
    // Cell formats are quoted.
    pub(crate) fn quote_string(&mut self) {
        // Only quote string values.
        if !self.is_string || self.value.is_empty() {
            return;
        }

        // Ignore already quoted strings.
        if self.value.starts_with('"') && self.value.ends_with('"') {
            return;
        }

        // Excel requires that double quotes are doubly quoted.
        self.value = self.value.replace('"', "\"\"");

        // Double quote the remaining string.
        self.value = format!("\"{}\"", self.value);
    }
}

macro_rules! conditional_format_value_from_string {
    ($($t:ty)*) => ($(
        impl From<$t> for ConditionalFormatValue {
            fn from(value: $t) -> ConditionalFormatValue {
                let mut value = ConditionalFormatValue::new_from_string(value);
                value.is_string = true;
                value
            }
        }
    )*)
}
conditional_format_value_from_string!(&str &String String Cow<'_, str>);

macro_rules! conditional_format_value_from_number {
    ($($t:ty)*) => ($(
        impl From<$t> for ConditionalFormatValue {
            fn from(value: $t) -> ConditionalFormatValue {
                ConditionalFormatValue::new_from_string(value.to_string())
            }
        }
    )*)
}
conditional_format_value_from_number!(u8 i8 u16 i16 u32 i32 f32 f64);

impl From<Formula> for ConditionalFormatValue {
    fn from(value: Formula) -> ConditionalFormatValue {
        ConditionalFormatValue::new_from_string(value.expand_formula(true))
    }
}

impl From<ExcelDateTime> for ConditionalFormatValue {
    fn from(value: ExcelDateTime) -> ConditionalFormatValue {
        let value = value.to_excel().to_string();
        ConditionalFormatValue::new_from_string(value)
    }
}

impl From<&ExcelDateTime> for ConditionalFormatValue {
    fn from(value: &ExcelDateTime) -> ConditionalFormatValue {
        let value = value.to_excel().to_string();
        ConditionalFormatValue::new_from_string(value)
    }
}

#[cfg(feature = "chrono")]
impl From<&NaiveDate> for ConditionalFormatValue {
    fn from(value: &NaiveDate) -> ConditionalFormatValue {
        let value = ExcelDateTime::chrono_date_to_excel(value).to_string();
        ConditionalFormatValue::new_from_string(value)
    }
}

#[cfg(feature = "chrono")]
impl From<&NaiveDateTime> for ConditionalFormatValue {
    fn from(value: &NaiveDateTime) -> ConditionalFormatValue {
        let value = ExcelDateTime::chrono_datetime_to_excel(value).to_string();
        ConditionalFormatValue::new_from_string(value)
    }
}

#[cfg(feature = "chrono")]
impl From<&NaiveTime> for ConditionalFormatValue {
    fn from(value: &NaiveTime) -> ConditionalFormatValue {
        let value = ExcelDateTime::chrono_time_to_excel(value).to_string();
        ConditionalFormatValue::new_from_string(value)
    }
}

// -----------------------------------------------------------------------
// ConditionalFormatCellCriteria
// -----------------------------------------------------------------------

/// The `ConditionalFormatCellCriteria` enum defines the conditional format
/// criteria for [`ConditionalFormatCell`] .
///
///
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ConditionalFormatCellCriteria {
    /// The cell conditional format criteria isn't set.
    #[doc(hidden)]
    None,

    /// Show the conditional format for cells that are equal to the target value.
    EqualTo,

    /// Show the conditional format for cells that are not equal to the target value.
    NotEqualTo,

    /// Show the conditional format for cells that are greater than the target value.
    GreaterThan,

    /// Show the conditional format for cells that are greater than or equal to the target value.
    GreaterThanOrEqualTo,

    /// Show the conditional format for cells that are less than the target value.
    LessThan,

    /// Show the conditional format for cells that are less than or equal to the target value.
    LessThanOrEqualTo,

    /// Show the conditional format for cells that are between the target values.
    Between,

    /// Show the conditional format for cells that are not between the target values.
    NotBetween,
}

impl fmt::Display for ConditionalFormatCellCriteria {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ConditionalFormatCellCriteria::None => write!(f, ""),
            ConditionalFormatCellCriteria::EqualTo => write!(f, "equal"),
            ConditionalFormatCellCriteria::Between => write!(f, "between"),
            ConditionalFormatCellCriteria::LessThan => write!(f, "lessThan"),
            ConditionalFormatCellCriteria::NotEqualTo => write!(f, "notEqual"),
            ConditionalFormatCellCriteria::NotBetween => write!(f, "notBetween"),
            ConditionalFormatCellCriteria::GreaterThan => write!(f, "greaterThan"),
            ConditionalFormatCellCriteria::LessThanOrEqualTo => write!(f, "lessThanOrEqual"),
            ConditionalFormatCellCriteria::GreaterThanOrEqualTo => write!(f, "greaterThanOrEqual"),
        }
    }
}

// -----------------------------------------------------------------------
// ConditionalFormatAverageCriteria
// -----------------------------------------------------------------------

/// The `ConditionalFormatAverageCriteria` enum defines the conditional format
/// criteria for [`ConditionalFormatCell`] .
///
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ConditionalFormatAverageCriteria {
    /// Show the conditional format for cells above the average for the range.
    /// This is the default.
    AboveAverage,

    /// Show the conditional format for cells below the average for the range.
    BelowAverage,

    /// Show the conditional format for cells above or equal to the average for
    /// the range.
    EqualOrAboveAverage,

    /// Show the conditional format for cells below or equal to the average for
    /// the range.
    EqualOrBelowAverage,

    /// Show the conditional format for cells 1 standard deviation above the
    /// average for the range.
    OneStandardDeviationAbove,

    /// Show the conditional format for cells 1 standard deviation below the
    /// average for the range.
    OneStandardDeviationBelow,

    /// Show the conditional format for cells 2 standard deviation above the
    /// average for the range.
    TwoStandardDeviationsAbove,

    /// Show the conditional format for cells 2 standard deviation below the
    /// average for the range.
    TwoStandardDeviationsBelow,

    /// Show the conditional format for cells 3 standard deviation above the
    /// average for the range.
    ThreeStandardDeviationsAbove,

    /// Show the conditional format for cells 3 standard deviation below the
    /// average for the range.
    ThreeStandardDeviationsBelow,
}

// -----------------------------------------------------------------------
// ConditionalFormatTextCriteria
// -----------------------------------------------------------------------

/// The `ConditionalFormatTextCriteria` enum defines the conditional format
/// criteria for [`ConditionalFormatText`] .
///
///
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ConditionalFormatTextCriteria {
    /// Show the conditional format for text that contains to the target string.
    Contains,

    /// Show the conditional format for text that do not contain to the target string.
    DoesNotContain,

    /// Show the conditional format for text that begins with the target string.
    BeginsWith,

    /// Show the conditional format for text that ends with the target string.
    EndsWith,
}

// -----------------------------------------------------------------------
// Generate common methods.
// -----------------------------------------------------------------------
macro_rules! generate_conditional_common_methods {
    ($($t:ty)*) => ($(

    /// **Section 2**: The following methods are common to all conditional
    /// formatting variants.
    impl $t {
        /// Set the [`Format`] of the conditional format rule.
        ///
        /// Set the [`Format`] that will be applied to the cell range if the conditional
        /// format rule applies. Not all cell format properties can be set in a
        /// conditional format. See [Excel's limitations on conditional format
        /// properties](crate::conditional_format#excels-limitations-on-conditional-format-properties) for
        /// more information.
        ///
        /// See the examples above.
        ///
        /// # Parameters
        ///
        /// * `format` - The [`Format`] property for the conditional format.
        ///
        pub fn set_format(mut self, format: impl Into<Format>) -> $t {
            self.format = Some(format.into());
            self
        }

        /// Set an additional multi-cell range for the conditional format.
        ///
        /// The `set_multi_range()` method is used to extend a conditional
        /// format over non-contiguous ranges like `"B3:D6 I3:K6 B9:D12
        /// I9:K12"`.
        ///
        /// See [Selecting a non-contiguous
        /// range](crate::conditional_format#selecting-a-non-contiguous-range)
        /// for more information.
        ///
        /// # Parameters
        ///
        /// * `range` - A string like type representing an Excel range.
        ///
        ///   Note, you can use an Excel range like `"$B$3:$D$6,$I$3:$K$6"` or
        ///   omit the `$` anchors and replace the commas with spaces to have a
        ///   clearer range like `"B3:D6 I3:K6"`. The documentation and examples
        ///   use the latter format for clarity but it you are copying and
        ///   pasting from Excel you can use the first format.
        ///
        pub fn set_multi_range(mut self, range: impl Into<String>) -> $t {
            self.multi_range = range.into().replace('$', "").replace(',', " ");
            self
        }

        /// Set the "Stop if True" option for the conditional format rule.
        ///
        /// The `set_stop_if_true()` method can be used to set the “Stop if true”
        /// feature of a conditional formatting rule when more than one rule is
        /// applied to a cell or a range of cells. When this parameter is set then
        /// subsequent rules are not evaluated if the current rule is true.
        ///
        /// # Parameters
        ///
        /// * `enable` - Turn the property on/off. It is off by default.
        ///
        pub fn set_stop_if_true(mut self, enable: bool) -> $t {
            self.stop_if_true = enable;
            self
        }

        // Get the index of the format object in the conditional format.
        pub(crate) fn format_index(&self) -> Option<u32> {
            self.format.as_ref().map(|format| format.dxf_index)
        }

        // Get a reference to the format object in the conditional format.
        pub(crate) fn format_as_mut(&mut self) -> Option<&mut Format> {
            self.format.as_mut()
        }

        // Get the multi-cell range for the conditional format, if present.
        pub(crate) fn multi_range(&self) -> String {
            self.multi_range.clone()
        }
    }
    )*)
}
generate_conditional_common_methods!(
    ConditionalFormatAverage
    ConditionalFormatCell
    ConditionalFormatDuplicate
    ConditionalFormatText
    ConditionalFormatTop
);
