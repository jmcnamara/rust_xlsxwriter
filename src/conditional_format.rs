// image - A module to represent Excel conditional formats.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! Working with Conditional Formats
//!
//! TODO
//!
//!
//!
//! # Excel's limitations on conditional format properties
//!
//! Not all of Excel's cell format properties can be modified with a conditional
//! format. Properties that **cannot** be modified in a conditional format are
//! font name, font size, superscript and subscript, diagonal borders, all
//! alignment properties and all protection properties.
//!
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
    fn get_rule_string(&self, dxf_index: Option<u32>, priority: u32) -> String;

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

            fn get_rule_string(&self, dxf_index: Option<u32>, priority: u32) -> String {
                self.get_rule_string(dxf_index, priority)
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
generate_conditional_format_impls!(ConditionalFormatCell ConditionalFormatDuplicate ConditionalFormatAverage);

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
    pub(crate) fn get_rule_string(&self, dxf_index: Option<u32>, priority: u32) -> String {
        let mut writer = XMLWriter::new();
        let mut attributes = vec![("type", "cellIs".to_string())];

        if let Some(dxf_index) = dxf_index {
            attributes.push(("dxfId", dxf_index.to_string()));
        }

        attributes.push(("priority", priority.to_string()));

        if self.stop_if_true {
            attributes.push(("stopIfTrue", "1".to_string()));
        }

        attributes.push(("operator", self.criteria.to_string()));

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
///     // Invert the duplicate conditional format to show uniques values.
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
    pub(crate) fn get_rule_string(&self, dxf_index: Option<u32>, priority: u32) -> String {
        let mut writer = XMLWriter::new();
        let mut attributes = vec![];

        if self.is_unique {
            attributes.push(("type", "uniqueValues".to_string()));
        } else {
            attributes.push(("type", "duplicateValues".to_string()));
        }

        if let Some(dxf_index) = dxf_index {
            attributes.push(("dxfId", dxf_index.to_string()));
        }

        attributes.push(("priority", priority.to_string()));

        if self.stop_if_true {
            attributes.push(("stopIfTrue", "1".to_string()));
        }

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
    pub(crate) fn get_rule_string(&self, dxf_index: Option<u32>, priority: u32) -> String {
        let mut writer = XMLWriter::new();
        let mut attributes = vec![("type", "aboveAverage".to_string())];

        if let Some(dxf_index) = dxf_index {
            attributes.push(("dxfId", dxf_index.to_string()));
        }

        attributes.push(("priority", priority.to_string()));

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

        writer.xml_empty_tag("cfRule", &attributes);

        writer.read_to_string()
    }
}

// -----------------------------------------------------------------------
// ConditionalFormatValue
// -----------------------------------------------------------------------

/// TODO
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
        /// properties](#excels-limitations-on-conditional-format-properties) for
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
        /// The `set_multi_range()` method is used to extend a conditional format
        /// over non-contiguous ranges.
        ///
        /// It is possible to apply a conditional format to different cell ranges in
        /// a worksheet using multiple calls to
        /// [`Worksheet::add_conditional_format()`](crate::Worksheet::add_conditional_format).
        /// However, as a minor optimization it is also possible in Excel to apply
        /// the same conditional format to different non-contiguous cell ranges.
        ///
        /// This is replicated in the `rust_xlsxwriter` conditional formats using
        /// the `set_multi_range()` method. The range must contain the primary range
        /// for the conditional format and any others separated by spaces. For
        /// example a range like `"A1:A3 A5 B3:K6 B9:K12"`.
        ///
        /// # Parameters
        ///
        /// * `range` - A string like type representing an Excel range.
        ///
        pub fn set_multi_range(mut self, range: impl Into<String>) -> $t {
            self.multi_range = range.into();
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
generate_conditional_common_methods!(ConditionalFormatCell ConditionalFormatDuplicate ConditionalFormatAverage);
