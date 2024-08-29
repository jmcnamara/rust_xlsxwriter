// data_validation - A module to represent Excel data validations.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

mod tests;

#[cfg(feature = "chrono")]
use chrono::{NaiveDate, NaiveDateTime, NaiveTime};

use crate::{ExcelDateTime, Formula, IntoExcelDateTime, XlsxError};
use std::fmt;

// -----------------------------------------------------------------------
// DataValidation
// -----------------------------------------------------------------------

/// The `DataValidation` struct represents a data validation in Excel.
///
/// `DataValidation` is used in conjunction with the
/// [`Worksheet::add_data_validation()`](crate::Worksheet::add_data_validation)
/// method.
///
/// # Working with Data Validation
///
/// Data validation is a feature of Excel that allows you to restrict the data
/// that a user enters in a cell and to display associated help and warning
/// messages. It also allows you to restrict input to values in a dropdown list.
///
/// A typical use case would be to restrict data in a cell to integer values
/// within a certain range, to provide a help message to explain the required
/// value, and to issue a warning if the input data doesn't meet the defined
/// criteria. For example:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/data_validation_intro1.png">
///
/// The example above was created with the following code:
///
/// ```
/// # // This code is available in examples/doc_data_validation_intro1.rs
/// #
/// use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///     let worksheet = workbook.add_worksheet();
///
///     worksheet.write(1, 0, "Enter rating in cell D2:")?;
///
///     let data_validation = DataValidation::new()
///         .allow_whole_number(DataValidationRule::Between(1, 5))
///         .set_input_title("Enter a star rating!")?
///         .set_input_message("Enter rating 1-5.\nWhole numbers only.")?
///         .set_error_title("Value outside allowed range")?
///         .set_error_message("The input value must be an integer in the range 1-5.")?;
///
///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
///
///     // Save the file.
///     workbook.save("data_validation.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Another common data validation task is to limit input to values on a
/// dropdown list.
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/data_validation_allow_list_strings.png">
///
/// The example above was created with the following code:
///
/// ```
/// # // This code is available in examples/doc_data_validation_allow_list_strings.rs
/// #
/// use rust_xlsxwriter::{DataValidation, Workbook, XlsxError};
///
/// fn main() -> Result<(), XlsxError> {
///     // Create a new Excel file object.
///     let mut workbook = Workbook::new();
///     let worksheet = workbook.add_worksheet();
///
///     worksheet.write(1, 0, "Select value in cell D2:")?;
///
///     let data_validation =
///         DataValidation::new().allow_list_strings(&["Pass", "Fail", "Incomplete"])?;
///
///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
///
///     // Save the file.
///     workbook.save("data_validation.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// For more information and examples see the API documentation below.
///
///
/// ## Using cell references in Data Validations
///
/// Excel allows the values for data validation to be either a "literal" number,
/// date or time value or a cell reference such as `=D1`. In the
/// `DataValidation` interfaces the cell reference values are represented using
/// a [`Formula`] value, like this:
///
/// ```
/// # // This code is available in examples/doc_data_validation_allow_whole_number_formula2.rs
/// #
/// # use rust_xlsxwriter::{DataValidation, DataValidationRule, Formula, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     worksheet.write(0, 0, "Upper limit:")?;
/// #     worksheet.write(0, 3, 10)?;
/// #     worksheet.write(1, 0, "Enter value in cell D2:")?;
/// #
///     let data_validation = DataValidation::new()
///         .allow_whole_number_formula(
///             DataValidationRule::LessThanOrEqualTo(Formula::new("=D1")));
/// #
/// #     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
/// #
/// #     // Save the file.
/// #     workbook.save("data_validation.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
/// In Excel this creates the following data validation:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/data_validation_allow_whole_number_formula.png">
///
/// As a syntactic shorthand you can also use `into()` for `Formula` like this:
///
/// ```
/// # // This code is available in examples/doc_data_validation_allow_whole_number_formula.rs
/// #
/// # use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     // Create a new Excel file object.
/// #     let mut workbook = Workbook::new();
/// #     let worksheet = workbook.add_worksheet();
/// #
/// #     worksheet.write(0, 0, "Upper limit:")?;
/// #     worksheet.write(0, 3, 10)?;
/// #     worksheet.write(1, 0, "Enter value in cell D2:")?;
/// #
///     let data_validation = DataValidation::new()
///         .allow_whole_number_formula(DataValidationRule::LessThanOrEqualTo("=D1".into()));
/// #
/// #     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
/// #
/// #     // Save the file.
/// #     workbook.save("data_validation.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// APIs that use [`Formula`] can also take Excel formulas, where appropriate.
///
/// **Note**: the contents of the `Formula` aren't validated by
/// `rust_xlsxwriter` so if you are using anything other than a simple cell
/// reference you should take care to ensure that the ranges or formulas are
/// valid in Excel.
///
#[derive(Clone)]
pub struct DataValidation {
    pub(crate) validation_type: DataValidationType,
    pub(crate) rule: DataValidationRuleInternal,
    pub(crate) ignore_blank: bool,
    pub(crate) show_input_message: bool,
    pub(crate) show_error_message: bool,
    pub(crate) show_dropdown: bool,
    pub(crate) multi_range: String,
    pub(crate) input_title: String,
    pub(crate) error_title: String,
    pub(crate) input_message: String,
    pub(crate) error_message: String,
    pub(crate) error_style: DataValidationErrorStyle,
}

impl DataValidation {
    /// Create a new cell Data Validation struct.
    ///
    /// The default type of a new data validation is equivalent to Excel's "Any"
    /// data validation. Refer to the `allow_TYPE()` functions below to
    /// constrain the data validation to defined types and to apply rules.
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> DataValidation {
        DataValidation {
            validation_type: DataValidationType::Any,
            rule: DataValidationRuleInternal::EqualTo(String::new()),
            ignore_blank: true,
            show_input_message: true,
            show_error_message: true,
            show_dropdown: true,
            multi_range: String::new(),
            input_title: String::new(),
            error_title: String::new(),
            input_message: String::new(),
            error_message: String::new(),
            error_style: DataValidationErrorStyle::Stop,
        }
    }

    /// Set a data validation to limit input to integers using defined rules.
    ///
    /// Set a data validation rule to restrict cell input to integers based on
    /// [`DataValidationRule`] rules such as "between" or "less than". Excel
    /// refers to this data validation type as "Whole number" but it equates to
    /// any Rust type that can convert [`Into`] a [`i32`] (`i64/u64` values
    /// aren't supported by Excel).
    ///
    /// # Parameters
    ///
    /// - `rule`: A [`DataValidationRule`] with [`i32`] values.
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// restricts input to integer values in a fixed range.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_allow_whole_number.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(1, 0, "Enter value in cell D2:")?;
    /// #
    ///     let data_validation =
    ///         DataValidation::new().allow_whole_number(DataValidationRule::Between(1, 10));
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_whole_number.png">
    ///
    /// The Excel Data Validation dialog for this file should look something like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_whole_number_dialog.png">
    ///
    pub fn allow_whole_number(mut self, rule: DataValidationRule<i32>) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Whole;
        self
    }

    /// Set a data validation to limit input to integers using defined rules and
    /// a cell reference.
    ///
    /// Set a data validation rule to restrict cell input to integers based on
    /// [`DataValidationRule`] rules such as "between" or "less than". Excel
    /// refers to this data validation type as "Whole number". The values used
    /// for the rules should be a cell reference represented by a [`Formula`],
    /// see [Using cell references in Data Validations] and the example below.
    ///
    /// # Parameters
    ///
    /// - `rule`: A [`DataValidationRule`] with cell references using
    ///   [`Formula`] values. See [Using cell references in Data Validations].
    ///
    /// [Using cell references in Data Validations]:
    ///     #using-cell-references-in-data-validations
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// restricts input to integer values based on a value from another cell.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_allow_whole_number_formula.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(0, 0, "Upper limit:")?;
    /// #     worksheet.write(0, 3, 10)?;
    /// #     worksheet.write(1, 0, "Enter value in cell D2:")?;
    /// #
    ///     let data_validation = DataValidation::new()
    ///         .allow_whole_number_formula(DataValidationRule::LessThanOrEqualTo("=D1".into()));
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// The Excel Data Validation dialog for the output file should look
    /// something like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_whole_number_formula.png">
    ///
    pub fn allow_whole_number_formula(
        mut self,
        rule: DataValidationRule<Formula>,
    ) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Whole;
        self
    }

    /// Set a data validation to limit input to floating point numbers using
    /// defined rules.
    ///
    /// Set a data validation rule to restrict cell input to floating point
    /// numbers based on [`DataValidationRule`] rules such as "between" or "less
    /// than". Excel refers to this data validation type as "Decimal" but it
    /// equates to the Rust type [`f64`].
    ///
    /// # Parameters
    ///
    /// - `rule`: A [`DataValidationRule`] with [`f64`] values.
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// restricts input to floating point values in a fixed range.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_allow_decimal_number.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(1, 0, "Enter value in cell D2:")?;
    /// #
    ///     let data_validation =
    ///         DataValidation::new().allow_decimal_number(DataValidationRule::Between(-9.9, 9.9));
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_decimal_number.png">
    ///
    /// The Excel Data Validation dialog for this file should look something like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_decimal_number_dialog.png">
    ///
    pub fn allow_decimal_number(mut self, rule: DataValidationRule<f64>) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Decimal;
        self
    }

    /// Set a data validation to limit input to floating point numbers using
    /// defined rules and a cell reference.
    ///
    /// Set a data validation rule to restrict cell input to floating point
    /// numbers based on [`DataValidationRule`] rules such as "between" or "less
    /// than". Excel refers to this data validation type as "Decimal". The
    /// values used for the rules should be a cell reference represented by a
    /// [`Formula`], see [Using cell references in Data Validations] and the
    /// example below.
    ///
    /// # Parameters
    ///
    /// - `rule`: A [`DataValidationRule`] with cell references using
    ///   [`Formula`] values. See [Using cell references in Data Validations].
    ///
    /// [Using cell references in Data Validations]:
    ///     #using-cell-references-in-data-validations
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// restricts input to floating point values based on a value from another
    /// cell.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_allow_decimal_number_formula.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(0, 0, "Upper limit:")?;
    /// #     worksheet.write(0, 3, 99.9)?;
    /// #     worksheet.write(1, 0, "Enter value in cell D2:")?;
    /// #
    ///     let data_validation = DataValidation::new()
    ///         .allow_decimal_number_formula(DataValidationRule::LessThanOrEqualTo("=D1".into()));
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// The Excel Data Validation dialog for the output file should look something like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_decimal_number_formula.png">
    ///
    pub fn allow_decimal_number_formula(
        mut self,
        rule: DataValidationRule<Formula>,
    ) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Decimal;
        self
    }

    /// Set a data validation rule to restrict cell input to a selection of
    /// strings via a dropdown menu.
    ///
    /// This type of validation presents the user with a list of allowed strings
    /// via a dropdown menu similar to online forms.
    ///
    /// Excel has a 255 character limit to the the string used to store the
    /// comma-separated list of strings, including the commas. This limit makes
    /// it unsuitable for long lists such as a list of provinces or states. For
    /// longer lists it is better to place the string values somewhere in the
    /// Excel workbook and refer to them using a range formula via the
    /// [`DataValidation::allow_list_formula()`] method shown below.
    ///
    /// # Parameters
    ///
    /// - `list`: A list of string like objects.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DataValidationError`] - The length of the combined
    ///   comma-separated list of strings, including commas, exceeds Excel's
    ///   limit of 255 characters, see the explanation above.
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// restricts users to a selection of values from a dropdown list.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_allow_list_strings.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(1, 0, "Select value in cell D2:")?;
    /// #
    ///     let data_validation =
    ///         DataValidation::new().allow_list_strings(&["Pass", "Fail", "Incomplete"])?;
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_list_strings.png">
    ///
    /// The Excel Data Validation dialog for this file should look something
    /// like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_list_strings_dialog.png">
    ///
    /// The following is a similar example but it demonstrates how to
    /// pre-populate a default choice.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_allow_list_strings2.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(1, 0, "Select value in cell D2:")?;
    /// #
    ///     let data_validation =
    ///         DataValidation::new().allow_list_strings(&["Pass", "Fail", "Incomplete"])?;
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    ///
    ///     // Add a default string to the cell with the data validation
    ///     // to pre-populate a default choice.
    ///     worksheet.write(1, 3, "Pass")?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    pub fn allow_list_strings(
        mut self,
        list: &[impl AsRef<str>],
    ) -> Result<DataValidation, XlsxError> {
        let joined_list = list
            .iter()
            .map(|s| s.as_ref().to_string().replace('"', "\"\""))
            .collect::<Vec<String>>()
            .join(",");

        let length = joined_list.chars().count();
        if length > 255 {
            return Err(XlsxError::DataValidationError(
                format!("Validation list length '{length}' including commas is greater than Excel's limit of 255 characters: {joined_list}")
            ));
        }

        let joined_list = format!("\"{joined_list}\"");

        self.rule = DataValidationRuleInternal::ListSource(joined_list);
        self.validation_type = DataValidationType::List;
        Ok(self)
    }

    /// Set a data validation rule to restrict cell input to a selection of
    /// strings via a dropdown menu and a cell range reference.
    ///
    /// The strings for the dropdown should be placed somewhere in the
    /// worksheet/workbook and should be referred to via a cell range reference
    /// represented by a [`Formula`], see [Using cell references in Data
    /// Validations] and the example below.
    ///
    /// # Parameters
    /// - `list`: A cell range reference such as `=B1:B9`, `=$B$1:$B$9` or
    ///   `=Sheet2!B1:B9` using a [`Formula`]. See [Using cell references in
    ///   Data Validations].
    ///
    /// [Using cell references in Data Validations]:
    ///     #using-cell-references-in-data-validations
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// restricts users to a selection of values from a dropdown list. The list data
    /// is provided from a cell range.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_allow_list_formula.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(1, 0, "Select value in cell D2:")?;
    /// #
    ///     // Write the string list data to some cells.
    ///     let string_list = ["Pass", "Fail", "Incomplete"];
    ///     worksheet.write_column(1, 5, string_list)?;
    ///
    ///     let data_validation = DataValidation::new().allow_list_formula("F2:F4".into());
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/data_validation_allow_list_formula.png">
    ///
    /// The Excel Data Validation dialog for this file should look something
    /// like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_list_formula_dialog.png">
    ///
    pub fn allow_list_formula(mut self, list: Formula) -> DataValidation {
        let formula = list.formula_string.clone();
        self.rule = DataValidationRuleInternal::ListSource(formula);
        self.validation_type = DataValidationType::List;
        self
    }

    /// Set a data validation to limit input to dates using defined rules.
    ///
    /// Set a data validation rule to restrict cell input to date values based
    /// on [`DataValidationRule`] rules such as "between" or "less than". Excel
    /// refers to this data validation type as "Date".
    ///
    /// This method uses date types that implement [`IntoExcelDateTime`]. The
    /// main date type supported is [`ExcelDateTime`]. If the `chrono` feature
    /// is enabled you can also use [`chrono::NaiveDate`].
    ///
    /// [`chrono::NaiveDate`]:
    ///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDate.html
    ///
    /// # Parameters
    ///
    /// - `rule`: A [`DataValidationRule`] with [`IntoExcelDateTime`] values.
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// restricts input to date values in a fixed range.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_allow_date.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, DataValidationRule, ExcelDateTime, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(1, 0, "Enter value in cell D2:")?;
    /// #
    ///     let data_validation = DataValidation::new().allow_date(DataValidationRule::Between(
    ///         ExcelDateTime::parse_from_str("2025-01-01")?,
    ///         ExcelDateTime::parse_from_str("2025-12-12")?,
    ///     ));
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// The Excel Data Validation dialog for this file should look something
    /// like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_date_dialog.png">
    ///
    pub fn allow_date(
        mut self,
        rule: DataValidationRule<impl IntoExcelDateTime + IntoDataValidationValue>,
    ) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Date;
        self
    }

    /// Set a data validation to limit input to dates using defined rules and a
    /// cell reference.
    ///
    /// Set a data validation rule to restrict cell input to date values based
    /// on [`DataValidationRule`] rules such as "between" or "less than". Excel
    /// refers to this data validation type as "Date".
    ///
    /// The values used for the rules should be a cell reference represented by
    /// a [`Formula`], see [Using cell references in Data Validations] and the
    /// `allow_TYPE_formula()` examples above.
    ///
    /// # Parameters
    ///
    /// - `rule`: A [`DataValidationRule`] with cell references using
    ///   [`Formula`] values. See [Using cell references in Data Validations].
    ///
    /// [Using cell references in Data Validations]:
    ///     #using-cell-references-in-data-validations
    ///
    pub fn allow_date_formula(mut self, rule: DataValidationRule<Formula>) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Date;
        self
    }

    /// Set a data validation to limit input to time values using defined rules.
    ///
    /// Set a data validation rule to restrict cell input to time values based
    /// on [`DataValidationRule`] rules such as "between" or "less than". Excel
    /// refers to this data validation type as "Time".
    ///
    /// This method uses time types that implement [`IntoExcelDateTime`]. The
    /// main time type supported is [`ExcelDateTime`]. If the `chrono` feature
    /// is enabled you can also use [`chrono::NaiveTime`].
    ///
    /// [`chrono::NaiveTime`]:
    ///     https://docs.rs/chrono/latest/chrono/naive/struct.NaiveTime.html
    ///
    /// # Parameters
    ///
    /// - `rule`: A [`DataValidationRule`] with [`IntoExcelDateTime`] values.
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// restricts input to time values in a fixed range.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_allow_time.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, DataValidationRule, ExcelDateTime, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(1, 0, "Enter value in cell D2:")?;
    /// #
    ///     let data_validation = DataValidation::new().allow_time(DataValidationRule::Between(
    ///         ExcelDateTime::parse_from_str("6:00")?,
    ///         ExcelDateTime::parse_from_str("12:00")?,
    ///     ));
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// The Excel Data Validation dialog for this file should look something
    /// like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_time_dialog.png">
    ///
    pub fn allow_time(
        mut self,
        rule: DataValidationRule<impl IntoExcelDateTime + IntoDataValidationValue>,
    ) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Time;
        self
    }

    /// Set a data validation to limit input to time values using defined rules
    /// and a cell reference.
    ///
    /// Set a data validation rule to restrict cell input to time values based
    /// on [`DataValidationRule`] rules such as "between" or "less than". Excel
    /// refers to this data validation type as "Time".
    ///
    /// The values used for the rules should be a cell reference represented by
    /// a [`Formula`], see [Using cell references in Data Validations] and the
    /// `allow_TYPE_formula()` examples above.
    ///
    /// # Parameters
    ///
    /// - `rule`: A [`DataValidationRule`] with cell references using
    ///   [`Formula`] values. See [Using cell references in Data Validations].
    ///
    /// [Using cell references in Data Validations]:
    ///     #using-cell-references-in-data-validations
    ///
    pub fn allow_time_formula(mut self, rule: DataValidationRule<Formula>) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::Time;
        self
    }

    /// Set a data validation to limit input to string lengths using defined
    /// rules.
    ///
    /// Set a data validation rule to restrict cell input to string lengths
    /// based on [`DataValidationRule`] rules such as "between" or "less than".
    /// Excel refers to this data validation type as "Text length" but it
    /// equates to any Rust type that can convert [`Into`] a [`u32`] (`u64`
    /// values or strings of that length aren't supported by Excel).
    ///
    /// # Parameters
    ///
    /// - `rule`: A [`DataValidationRule`] with [`u32`] values.
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// restricts input to strings whose length is in a fixed range.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_allow_text_length.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(1, 0, "Enter value in cell D2:")?;
    /// #
    ///     let data_validation =
    ///         DataValidation::new().allow_text_length(DataValidationRule::Between(4, 8));
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// The Excel Data Validation dialog for this file should look something
    /// like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_text_length_dialog.png">
    ///
    pub fn allow_text_length(mut self, rule: DataValidationRule<u32>) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::TextLength;
        self
    }

    /// Set a data validation to limit input to string lengths using defined
    /// rules and a cell reference.
    ///
    /// Set a data validation rule to restrict cell input to string lengths
    /// based on [`DataValidationRule`] rules such as "between" or "less than".
    /// Excel refers to this data validation type as "Text length".
    ///
    /// The values used for the rules should be a cell reference represented by
    /// a [`Formula`], see [Using cell references in Data Validations] and the
    /// `allow_TYPE_formula()` examples above.
    ///
    /// # Parameters
    ///
    /// - `rule`: A [`DataValidationRule`] with cell references using
    ///   [`Formula`] values. See [Using cell references in Data Validations].
    ///
    /// [Using cell references in Data Validations]:
    ///     #using-cell-references-in-data-validations
    ///
    pub fn allow_text_length_formula(
        mut self,
        rule: DataValidationRule<Formula>,
    ) -> DataValidation {
        self.rule = rule.to_internal_rule();
        self.validation_type = DataValidationType::TextLength;
        self
    }

    /// Set a data validation to limit input based on a custom formula.
    ///
    /// Set a data validation rule to restrict cell input based on an Excel
    /// formula that returns a boolean value. Excel refers to this data
    /// validation type as "Custom".
    ///
    /// # Parameters
    ///
    /// - `rule`: A [`Formula`] value. You should ensure that the formula is
    ///   valid in Excel.
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// restricts input to text/strings that are uppercase.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_allow_custom.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(1, 0, "Enter uppercase string in D2:")?;
    /// #
    ///     let data_validation =
    ///         DataValidation::new().allow_custom("=AND(ISTEXT(D2), EXACT(D2, UPPER(D2)))".into());
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// The Excel Data Validation dialog for this file should look something
    /// like this:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_custom_dialog.png">
    ///
    pub fn allow_custom(mut self, rule: Formula) -> DataValidation {
        let formula = rule.formula_string.clone();
        self.rule = DataValidationRuleInternal::CustomFormula(formula);
        self.validation_type = DataValidationType::Custom;
        self
    }

    /// Set a data validation to allow any input data.
    ///
    /// The "Any" data validation type doesn't restrict data input and is mainly
    /// used to allow access to the "Input Message" dialog when a user enters data
    /// in a cell.
    ///
    /// This is the default validation type for [`DataValidation`] if no other
    /// `allow_TYPE()` method is used. Situations where this type of data
    /// validation are required are uncommon.
    ///
    pub fn allow_any_value(mut self) -> DataValidation {
        self.rule = DataValidationRuleInternal::EqualTo(String::new());
        self.validation_type = DataValidationType::Any;
        self
    }

    /// Set the data validation option that defines how blank cells are handled.
    ///
    /// By default Excel data validations have an "Ignore blank" option turned
    /// on. This allows the user to optionally leave the cell blank and not
    /// enter any value. This is generally the best default option since it
    /// allows users to exit the cell without inputting any data.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_whole_number_dialog.png">
    ///
    /// If you need to ensure that the user inserts some information then you
    /// can use `ignore_blank()` to turn this option off.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    pub fn ignore_blank(mut self, enable: bool) -> DataValidation {
        self.ignore_blank = enable;
        self
    }

    /// Turn on/off the in-cell dropdown for list data validations.
    ///
    /// By default the Excel list data validation has an "In-cell drop-down"
    /// option turned on. This shows a dropdown arrow for list style data
    /// validations and displays the list items.
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_allow_list_strings_dialog.png">
    ///
    /// If this option is turned off the data validation will restrict input to
    /// the specified list values but it won't display a visual indicator of
    /// what those values are.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    pub fn show_dropdown(mut self, enable: bool) -> DataValidation {
        self.show_dropdown = enable;
        self
    }

    /// Toggle option to show an input message when the cell is entered.
    ///
    /// This function is used to toggle the option that controls whether an
    /// input message is shown when a data validation cell is entered.
    ///
    /// The option only has an effect if there is an input message, so for the
    /// majority of use cases it isn't required.
    ///
    /// See also [`DataValidation::set_input_message()`] below.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    pub fn show_input_message(mut self, enable: bool) -> DataValidation {
        self.show_input_message = enable;
        self
    }

    /// Set the title for the input message when the cell is entered.
    ///
    /// This option is used to set a title in bold for the input message when a
    /// data validation cell is entered.
    ///
    /// The title is only visible if there is also an input message. See the
    /// [`DataValidation::set_input_message()`] example below.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DataValidationError`] - The length of the title exceeds
    ///   Excel's limit of 32 characters.
    ///
    /// # Parameters
    ///
    /// - `text`: Title string. Must be less than or equal to the Excel limit
    ///   of 32 characters.
    ///
    pub fn set_input_title(mut self, text: impl Into<String>) -> Result<DataValidation, XlsxError> {
        let text = text.into();
        let length = text.chars().count();

        if length > 32 {
            return Err(XlsxError::DataValidationError(format!(
                "Validation title length '{length}' greater than Excel's limit of 32 characters."
            )));
        }

        self.input_title = text;
        Ok(self)
    }

    /// Set the input message when a data validation cell is entered.
    ///
    /// This option is used to set an input message when a data validation cell
    /// is entered. This can we used to explain to the user what the data
    /// validation rules are for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DataValidationError`] - The length of the message exceeds
    ///   Excel's limit of 255 characters.
    ///
    /// # Parameters
    ///
    /// - `text`: Message string. Must be less than or equal to the Excel limit
    ///   of 255 characters. The string can contain newlines to split it over
    ///   several lines.
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// uses an input message to explain to the user what type of input is
    /// required.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_set_input_message.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(1, 0, "Enter rating in cell D2:")?;
    /// #
    ///     let data_validation = DataValidation::new()
    ///         .allow_whole_number(DataValidationRule::Between(1, 5))
    ///         .set_input_title("Enter a star rating!")?
    ///         .set_input_message("Enter rating 1-5.\nWhole numbers only.")?;
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_set_input_message.png">
    ///
    pub fn set_input_message(
        mut self,
        text: impl Into<String>,
    ) -> Result<DataValidation, XlsxError> {
        let text = text.into();
        let length = text.chars().count();

        if length > 255 {
            return Err(XlsxError::DataValidationError(format!(
                    "Validation message length '{length}' greater than Excel's limit of 255 characters."
            )));
        }

        self.input_message = text;
        Ok(self)
    }

    /// Toggle option to show an error message when there is a validation error.
    ///
    /// This function is used to toggle the option that controls whether an
    /// error message is shown when there is a validation error.
    ///
    /// If this option is toggled off then any data can be entered in a cell and
    /// an error message will not be raised, which has limited practical
    /// applications for a data validation.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    pub fn show_error_message(mut self, enable: bool) -> DataValidation {
        self.show_error_message = enable;
        self
    }

    /// Set the title for the error message when there is a validation error.
    ///
    /// This option is used to set a title in bold for the error message when
    /// there is a validation error.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DataValidationError`] - The length of the title exceeds
    ///   Excel's limit of 32 characters.
    ///
    /// # Parameters
    ///
    /// - `text`: Title string. Must be less than or equal to the Excel limit
    ///   of 32 characters.
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// shows a custom error title.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_set_error_title.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(1, 0, "Enter value in cell D2:")?;
    /// #
    ///     let data_validation = DataValidation::new()
    ///         .allow_whole_number(DataValidationRule::Between(1, 10))
    ///         .set_error_title("Danger, Will Robinson!")?;
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/data_validation_set_error_title.png">
    ///
    pub fn set_error_title(mut self, text: impl Into<String>) -> Result<DataValidation, XlsxError> {
        let text = text.into();
        let length = text.chars().count();

        if length > 32 {
            return Err(XlsxError::DataValidationError(format!(
                "Validation title length '{length}' greater than Excel's limit of 32 characters."
            )));
        }

        self.error_title = text;
        Ok(self)
    }

    /// Set the error message when there is a validation error.
    ///
    /// This option is used to set an error message when there is a validation
    /// error. This can we used to explain to the user what the data validation
    /// rules are for the cell.
    ///
    /// # Errors
    ///
    /// - [`XlsxError::DataValidationError`] - The length of the message exceeds
    ///   Excel's limit of 255 characters.
    ///
    /// # Parameters
    ///
    /// - `text`: Message string. Must be less than or equal to the Excel limit
    ///   of 255 characters. The string can contain newlines to split it over
    ///   several lines.
    ///
    /// # Examples
    ///
    /// Example of adding a data validation to a worksheet cell. This validation
    /// shows a custom error message.
    ///
    /// ```
    /// # // This code is available in examples/doc_data_validation_set_error_message.rs
    /// #
    /// # use rust_xlsxwriter::{DataValidation, DataValidationRule, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     // Create a new Excel file object.
    /// #     let mut workbook = Workbook::new();
    /// #     let worksheet = workbook.add_worksheet();
    /// #
    /// #     worksheet.write(1, 0, "Enter value in cell D2:")?;
    /// #
    ///     let data_validation = DataValidation::new()
    ///         .allow_whole_number(DataValidationRule::Between(1, 10))
    ///         .set_error_title("Value outside allowed range")?
    ///         .set_error_message("The input value must be an integer in the range 1-10.")?;
    ///
    ///     worksheet.add_data_validation(1, 3, 1, 3, &data_validation)?;
    /// #
    /// #     // Save the file.
    /// #     workbook.save("data_validation.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/data_validation_set_error_message.png">
    ///
    pub fn set_error_message(
        mut self,
        text: impl Into<String>,
    ) -> Result<DataValidation, XlsxError> {
        let text = text.into();
        let length = text.chars().count();

        if length > 255 {
            return Err(XlsxError::DataValidationError(format!(
                    "Validation message length '{length}' greater than Excel's limit of 255 characters."
            )));
        }

        self.error_message = text;
        Ok(self)
    }

    /// Set the style of the error dialog type for a data validation.
    ///
    /// Set the error dialog to be either "Stop" (the default), "Warning" or
    /// "Information". This option only has an effect on Windows.
    ///
    /// # Parameters
    ///
    /// - `error_style`: A [`DataValidationErrorStyle`] enum value.
    ///
    pub fn set_error_style(mut self, error_style: DataValidationErrorStyle) -> DataValidation {
        self.error_style = error_style;
        self
    }

    /// Set an additional multi-cell range for the data validation.
    ///
    /// The `set_multi_range()` method is used to extend a data validation
    /// over non-contiguous ranges like `"B3 I3 B9:D12 I9:K12"`.
    ///
    /// # Parameters
    ///
    /// - `range`: A string like type representing an Excel range.
    ///
    pub fn set_multi_range(mut self, range: impl Into<String>) -> DataValidation {
        self.multi_range = range.into().replace('$', "").replace(',', " ");
        self
    }

    // -----------------------------------------------------------------------
    // Crate level helper methods.
    // -----------------------------------------------------------------------

    // The "Any" validation type should be ignored if it doesn't have any input
    // or error titles or messages. This is the same rule as Excel.
    pub(crate) fn is_invalid_any(&mut self) -> bool {
        self.validation_type == DataValidationType::Any
            && self.input_title.is_empty()
            && self.input_message.is_empty()
            && self.error_title.is_empty()
            && self.error_message.is_empty()
    }
}

/// Trait to map rust types into data validation types
///
/// The `IntoDataValidationValue` trait is used to map Rust types like numbers,
/// dates, times and formulas into a generic type that can be used to replicate
/// Excel data types used in Data Validation.
///
pub trait IntoDataValidationValue {
    /// Function to turn types into a string value.
    fn to_string_value(&self) -> String;
}

impl IntoDataValidationValue for i32 {
    fn to_string_value(&self) -> String {
        self.to_string()
    }
}

impl IntoDataValidationValue for u32 {
    fn to_string_value(&self) -> String {
        self.to_string()
    }
}

impl IntoDataValidationValue for f64 {
    fn to_string_value(&self) -> String {
        self.to_string()
    }
}

impl IntoDataValidationValue for Formula {
    fn to_string_value(&self) -> String {
        self.formula_string.clone()
    }
}

impl IntoDataValidationValue for ExcelDateTime {
    fn to_string_value(&self) -> String {
        self.to_excel().to_string()
    }
}

impl IntoDataValidationValue for &ExcelDateTime {
    fn to_string_value(&self) -> String {
        self.to_excel().to_string()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoDataValidationValue for NaiveDateTime {
    fn to_string_value(&self) -> String {
        ExcelDateTime::chrono_datetime_to_excel(self).to_string()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoDataValidationValue for &NaiveDateTime {
    fn to_string_value(&self) -> String {
        ExcelDateTime::chrono_datetime_to_excel(self).to_string()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoDataValidationValue for NaiveDate {
    fn to_string_value(&self) -> String {
        ExcelDateTime::chrono_date_to_excel(self).to_string()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoDataValidationValue for &NaiveDate {
    fn to_string_value(&self) -> String {
        ExcelDateTime::chrono_date_to_excel(self).to_string()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoDataValidationValue for NaiveTime {
    fn to_string_value(&self) -> String {
        ExcelDateTime::chrono_time_to_excel(self).to_string()
    }
}

#[cfg(feature = "chrono")]
#[cfg_attr(docsrs, doc(cfg(feature = "chrono")))]
impl IntoDataValidationValue for &NaiveTime {
    fn to_string_value(&self) -> String {
        ExcelDateTime::chrono_time_to_excel(self).to_string()
    }
}

// -----------------------------------------------------------------------
// DataValidationType
// -----------------------------------------------------------------------

/// The `DataValidationType` enum defines the type of data validation.
#[derive(Clone, Eq, PartialEq)]
pub(crate) enum DataValidationType {
    Whole,

    Decimal,

    Date,

    Time,

    TextLength,

    Custom,

    List,

    Any,
}

impl fmt::Display for DataValidationType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Any => write!(f, "any"),
            Self::Date => write!(f, "date"),
            Self::List => write!(f, "list"),
            Self::Time => write!(f, "time"),
            Self::Whole => write!(f, "whole"),
            Self::Custom => write!(f, "custom"),
            Self::Decimal => write!(f, "decimal"),
            Self::TextLength => write!(f, "textLength"),
        }
    }
}

// -----------------------------------------------------------------------
// DataValidationRule
// -----------------------------------------------------------------------

/// The `DataValidationRule` enum defines the data validation rule for
/// [`DataValidation`].
///
#[derive(Clone)]
pub enum DataValidationRule<T: IntoDataValidationValue> {
    /// Restrict cell input to values that are equal to the target value.
    EqualTo(T),

    /// Restrict cell input to values that are not equal to the target value.
    NotEqualTo(T),

    /// Restrict cell input to values that are greater than the target value.
    GreaterThan(T),

    /// Restrict cell input to values that are greater than or equal to the
    /// target value.
    GreaterThanOrEqualTo(T),

    /// Restrict cell input to values that are less than the target value.
    LessThan(T),

    /// Restrict cell input to values that are less than or equal to the target
    /// value.
    LessThanOrEqualTo(T),

    /// Restrict cell input to values that are between the target values.
    Between(T, T),

    /// Restrict cell input to values that are not between the target values.
    NotBetween(T, T),
}

impl<T: IntoDataValidationValue> DataValidationRule<T> {
    fn to_internal_rule(&self) -> DataValidationRuleInternal {
        match &self {
            DataValidationRule::EqualTo(value) => {
                DataValidationRuleInternal::EqualTo(value.to_string_value())
            }
            DataValidationRule::NotEqualTo(value) => {
                DataValidationRuleInternal::NotEqualTo(value.to_string_value())
            }
            DataValidationRule::GreaterThan(value) => {
                DataValidationRuleInternal::GreaterThan(value.to_string_value())
            }

            DataValidationRule::GreaterThanOrEqualTo(value) => {
                DataValidationRuleInternal::GreaterThanOrEqualTo(value.to_string_value())
            }
            DataValidationRule::LessThan(value) => {
                DataValidationRuleInternal::LessThan(value.to_string_value())
            }
            DataValidationRule::LessThanOrEqualTo(value) => {
                DataValidationRuleInternal::LessThanOrEqualTo(value.to_string_value())
            }
            DataValidationRule::Between(min, max) => {
                DataValidationRuleInternal::Between(min.to_string_value(), max.to_string_value())
            }
            DataValidationRule::NotBetween(min, max) => {
                DataValidationRuleInternal::NotBetween(min.to_string_value(), max.to_string_value())
            }
        }
    }
}

// -----------------------------------------------------------------------
// DataValidationRuleInternal
// -----------------------------------------------------------------------

// This is a variation on `DataValidationRule` that is used for internal storage
// of the validation rule. It only uses the String type.
#[derive(Clone)]
pub(crate) enum DataValidationRuleInternal {
    EqualTo(String),

    NotEqualTo(String),

    GreaterThan(String),

    GreaterThanOrEqualTo(String),

    LessThan(String),

    LessThanOrEqualTo(String),

    Between(String, String),

    NotBetween(String, String),

    CustomFormula(String),

    ListSource(String),
}

impl fmt::Display for DataValidationRuleInternal {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::EqualTo(_) => write!(f, "equal"),
            Self::LessThan(_) => write!(f, "lessThan"),
            Self::Between(_, _) => write!(f, "between"),
            Self::ListSource(_) => write!(f, "list"),
            Self::NotEqualTo(_) => write!(f, "notEqual"),
            Self::GreaterThan(_) => write!(f, "greaterThan"),
            Self::CustomFormula(_) => write!(f, ""),
            Self::NotBetween(_, _) => write!(f, "notBetween"),
            Self::LessThanOrEqualTo(_) => write!(f, "lessThanOrEqual"),
            Self::GreaterThanOrEqualTo(_) => write!(f, "greaterThanOrEqual"),
        }
    }
}

// -----------------------------------------------------------------------
// DataValidationErrorStyle
// -----------------------------------------------------------------------

/// The `DataValidationErrorStyle` enum defines the type of error dialog that is
/// shown when there is and error in a data validation.
///
#[derive(Clone)]
pub enum DataValidationErrorStyle {
    /// Show a "Stop" dialog. This is the default.
    Stop,

    /// Show a "Warning" dialog.
    Warning,

    /// Show an "Information" dialog.
    Information,
}

impl fmt::Display for DataValidationErrorStyle {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Stop => write!(f, "stop"),
            Self::Warning => write!(f, "warning"),
            Self::Information => write!(f, "information"),
        }
    }
}
