// image - A module to represent Excel conditional formats.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

mod tests;

#[cfg(feature = "chrono")]
use chrono::{NaiveDate, NaiveDateTime, NaiveTime};

use crate::{xmlwriter::XMLWriter, ExcelDateTime, Format, Formula, XlsxError};
use std::{borrow::Cow, fmt};

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
    stop_if_true: bool,
    pub(crate) format: Option<Format>,
}

impl ConditionalFormatCell {
    /// Create a new Cell conditional format struct.
    #[allow(clippy::new_without_default)]
    pub fn new() -> ConditionalFormatCell {
        ConditionalFormatCell {
            minimum: ConditionalFormatValue::new_from_string(""),
            maximum: ConditionalFormatValue::new_from_string(""),
            criteria: ConditionalFormatCellCriteria::None,
            stop_if_true: false,
            format: None,
        }
    }

    /// Set the value of the Cell conditional format rule.
    pub fn set_value(self, value: impl Into<ConditionalFormatValue>) -> ConditionalFormatCell {
        self.set_minimum(value)
    }

    /// Set the minimum value of the Cell conditional format rule.
    pub fn set_minimum(
        mut self,
        value: impl Into<ConditionalFormatValue>,
    ) -> ConditionalFormatCell {
        self.minimum = value.into();
        self.minimum.quote_string();
        self
    }

    /// Set the maximum value of the Cell conditional format rule.
    pub fn set_maximum(
        mut self,
        value: impl Into<ConditionalFormatValue>,
    ) -> ConditionalFormatCell {
        self.maximum = value.into();
        self.maximum.quote_string();
        self
    }

    /// Set the criteria of the Cell conditional format rule.
    pub fn set_criteria(
        mut self,
        criteria: ConditionalFormatCellCriteria,
    ) -> ConditionalFormatCell {
        self.criteria = criteria;
        self
    }

    /// Set the format of the Cell conditional format rule.
    pub fn set_format(mut self, format: impl Into<Format>) -> ConditionalFormatCell {
        self.format = Some(format.into());
        self
    }

    /// Set the "Stop if True" option for the Cell conditional format rule.
    pub fn set_stop_if_true(mut self, enable: bool) -> ConditionalFormatCell {
        self.stop_if_true = enable;
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
