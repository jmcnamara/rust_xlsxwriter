// filter - A module for representing autofilter conditions.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

/// The `FilterCondition` struct is used to define autofilter rules.
///
/// Autofilter rules are associated with ranges created using
/// [`autofilter()`](crate::Worksheet::autofilter()).
///
/// Excel supports two main types of filter conditions. The first, and most
/// common, is a list filter where the user selects the items to filter from a
/// list of all the values in the the column range:
///
/// <img src="https://rustxlsxwriter.github.io/images/autofilter_list.png">
///
/// The other main type of filter is a custom filter where the user can specify
/// 1 or 2 conditions like ">= 4000" and "<= 6000":
///
/// <img src="https://rustxlsxwriter.github.io/images/autofilter_custom.png">
///
/// In Excel these are mutually exclusive and you will need to choose one or the
/// other via the [`FilterCondition`]
/// [`add_list_filter()`](FilterCondition::add_list_filter) and
/// [`add_custom_filter()`](FilterCondition::add_custom_filter) methods.
///
///
///
/// # Examples
///
/// The following example demonstrates setting an autofilter with a list filter
/// condition.
///
/// ```
/// # // This code is available in examples/doc_worksheet_filter_column1.rs
/// #
/// # use rust_xlsxwriter::{FilterCondition, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet with some sample data to filter.
/// #     let worksheet = workbook.add_worksheet();
/// #     worksheet.write_string(0, 0, "Region")?;
/// #     worksheet.write_string(1, 0, "East")?;
/// #     worksheet.write_string(2, 0, "West")?;
/// #     worksheet.write_string(3, 0, "East")?;
/// #     worksheet.write_string(4, 0, "North")?;
/// #     worksheet.write_string(5, 0, "South")?;
/// #     worksheet.write_string(6, 0, "West")?;
/// #
/// #     worksheet.write_string(0, 1, "Sales")?;
/// #     worksheet.write_number(1, 1, 3000)?;
/// #     worksheet.write_number(2, 1, 8000)?;
/// #     worksheet.write_number(3, 1, 5000)?;
/// #     worksheet.write_number(4, 1, 4000)?;
/// #     worksheet.write_number(5, 1, 7000)?;
/// #     worksheet.write_number(6, 1, 9000)?;
/// #
/// #     // Set the autofilter.
/// #     worksheet.autofilter(0, 0, 6, 1)?;
///
///     // Set a filter condition to only show cells matching "East" in the first
///     // column.
///     let filter_condition = FilterCondition::new().add_list_filter("East");
///     worksheet.filter_column(0, &filter_condition)?;
///
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/worksheet_filter_column1.png">
///
///
/// The following example demonstrates setting an autofilter with multiple list
/// filter conditions.
///
/// ```
/// # // This code is available in examples/doc_worksheet_filter_column2.rs
/// #
/// # use rust_xlsxwriter::{FilterCondition, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet with some sample data to filter.
/// #     let worksheet = workbook.add_worksheet();
/// #     worksheet.write_string(0, 0, "Region")?;
/// #     worksheet.write_string(1, 0, "East")?;
/// #     worksheet.write_string(2, 0, "West")?;
/// #     worksheet.write_string(3, 0, "East")?;
/// #     worksheet.write_string(4, 0, "North")?;
/// #     worksheet.write_string(5, 0, "South")?;
/// #     worksheet.write_string(6, 0, "West")?;
/// #
/// #     worksheet.write_string(0, 1, "Sales")?;
/// #     worksheet.write_number(1, 1, 3000)?;
/// #     worksheet.write_number(2, 1, 8000)?;
/// #     worksheet.write_number(3, 1, 5000)?;
/// #     worksheet.write_number(4, 1, 4000)?;
/// #     worksheet.write_number(5, 1, 7000)?;
/// #     worksheet.write_number(6, 1, 9000)?;
/// #
/// #     // Set the autofilter.
/// #     worksheet.autofilter(0, 0, 6, 1)?;
///
///     // Set a filter condition to only show cells matching "East", "West" or
///     // "South" in the first column.
///     let filter_condition = FilterCondition::new()
///         .add_list_filter("East")
///         .add_list_filter("West")
///         .add_list_filter("South");
///     worksheet.filter_column(0, &filter_condition)?;
///
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/worksheet_filter_column2.png">
///
///
/// The following example demonstrates setting an autofilter with a list filter
/// for blank cells.
///
/// ```
/// # // This code is available in examples/doc_worksheet_filter_column3.rs
/// #
/// # use rust_xlsxwriter::{FilterCondition, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet with some sample data to filter.
/// #     let worksheet = workbook.add_worksheet();
/// #     worksheet.write_string(0, 0, "Region")?;
/// #     worksheet.write_string(1, 0, "")?;
/// #     worksheet.write_string(2, 0, "West")?;
/// #     worksheet.write_string(3, 0, "East")?;
/// #     worksheet.write_string(4, 0, "North")?;
/// #
/// #     worksheet.write_string(6, 0, "West")?;
/// #
/// #     worksheet.write_string(0, 1, "Sales")?;
/// #     worksheet.write_number(1, 1, 3000)?;
/// #     worksheet.write_number(2, 1, 8000)?;
/// #     worksheet.write_number(3, 1, 5000)?;
/// #     worksheet.write_number(4, 1, 4000)?;
/// #     worksheet.write_number(5, 1, 7000)?;
/// #     worksheet.write_number(6, 1, 9000)?;
/// #
/// #     // Set the autofilter.
/// #     worksheet.autofilter(0, 0, 6, 1)?;
///
///     // Set a filter condition to only show cells matching blanks.
///     let filter_condition = FilterCondition::new().add_list_blanks_filter();
///
///     worksheet.filter_column(0, &filter_condition)?;
///
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/worksheet_filter_column3.png">
///
///
/// The following example demonstrates setting an autofilter with different list
/// filter conditions in separate columns.
///
/// ```
/// # // This code is available in examples/doc_worksheet_filter_column4.rs
/// #
/// # use rust_xlsxwriter::{FilterCondition, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet with some sample data to filter.
/// #     let worksheet = workbook.add_worksheet();
/// #     worksheet.write_string(0, 0, "Region")?;
/// #     worksheet.write_string(1, 0, "East")?;
/// #     worksheet.write_string(2, 0, "West")?;
/// #     worksheet.write_string(3, 0, "East")?;
/// #     worksheet.write_string(4, 0, "North")?;
/// #     worksheet.write_string(5, 0, "South")?;
/// #     worksheet.write_string(6, 0, "West")?;
/// #
/// #     worksheet.write_string(0, 1, "Sales")?;
/// #     worksheet.write_number(1, 1, 3000)?;
/// #     worksheet.write_number(2, 1, 8000)?;
/// #     worksheet.write_number(3, 1, 5000)?;
/// #     worksheet.write_number(4, 1, 4000)?;
/// #     worksheet.write_number(5, 1, 7000)?;
/// #     worksheet.write_number(6, 1, 9000)?;
/// #
/// #     // Set the autofilter.
/// #     worksheet.autofilter(0, 0, 6, 1)?;
///
///     // Set a filter condition for 2 separate columns.
///     let filter_condition1 = FilterCondition::new().add_list_filter("East");
///     worksheet.filter_column(0, &filter_condition1)?;
///
///     let filter_condition2 = FilterCondition::new().add_list_filter(3000);
///     worksheet.filter_column(1, &filter_condition2)?;
///
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/worksheet_filter_column4.png">
///
///
/// The following example demonstrates setting an autofilter for a custom number
/// filter.
///
/// ```
/// # // This code is available in examples/doc_worksheet_filter_column5.rs
/// #
/// # use rust_xlsxwriter::{FilterCondition, FilterCriteria, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet with some sample data to filter.
/// #     let worksheet = workbook.add_worksheet();
/// #     worksheet.write_string(0, 0, "Region")?;
/// #     worksheet.write_string(1, 0, "East")?;
/// #     worksheet.write_string(2, 0, "West")?;
/// #     worksheet.write_string(3, 0, "East")?;
/// #     worksheet.write_string(4, 0, "North")?;
/// #     worksheet.write_string(5, 0, "South")?;
/// #     worksheet.write_string(6, 0, "West")?;
/// #
/// #     worksheet.write_string(0, 1, "Sales")?;
/// #     worksheet.write_number(1, 1, 3000)?;
/// #     worksheet.write_number(2, 1, 8000)?;
/// #     worksheet.write_number(3, 1, 5000)?;
/// #     worksheet.write_number(4, 1, 4000)?;
/// #     worksheet.write_number(5, 1, 7000)?;
/// #     worksheet.write_number(6, 1, 9000)?;
/// #
/// #     // Set the autofilter.
/// #     worksheet.autofilter(0, 0, 6, 1)?;
///
///     // Set a custom number filter.
///     let filter_condition =
///         FilterCondition::new().add_custom_filter(FilterCriteria::GreaterThan, 4000);
///     worksheet.filter_column(1, &filter_condition)?;
///
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/worksheet_filter_column5.png">
///
///
/// The following example demonstrates setting an autofilter for two custom
/// number filters to create a "between" condition.
///
/// ```
/// # // This code is available in examples/doc_worksheet_filter_column6.rs
/// #
/// # use rust_xlsxwriter::{FilterCondition, FilterCriteria, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet with some sample data to filter.
/// #     let worksheet = workbook.add_worksheet();
/// #     worksheet.write_string(0, 0, "Region")?;
/// #     worksheet.write_string(1, 0, "East")?;
/// #     worksheet.write_string(2, 0, "West")?;
/// #     worksheet.write_string(3, 0, "East")?;
/// #     worksheet.write_string(4, 0, "North")?;
/// #     worksheet.write_string(5, 0, "South")?;
/// #     worksheet.write_string(6, 0, "West")?;
/// #
/// #     worksheet.write_string(0, 1, "Sales")?;
/// #     worksheet.write_number(1, 1, 3000)?;
/// #     worksheet.write_number(2, 1, 8000)?;
/// #     worksheet.write_number(3, 1, 5000)?;
/// #     worksheet.write_number(4, 1, 4000)?;
/// #     worksheet.write_number(5, 1, 7000)?;
/// #     worksheet.write_number(6, 1, 9000)?;
/// #
/// #     // Set the autofilter.
/// #     worksheet.autofilter(0, 0, 6, 1)?;
///
///     // Set two custom number filters in a "between" configuration.
///     let filter_condition = FilterCondition::new()
///         .add_custom_filter(FilterCriteria::GreaterThanOrEqualTo, 4000)
///         .add_custom_filter(FilterCriteria::LessThanOrEqualTo, 8000);
///     worksheet.filter_column(1, &filter_condition)?;
///
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/worksheet_filter_column6.png">
///
///
/// The following example demonstrates setting an autofilter to show all the
/// non-blank values in a column. This can be done in 2 ways: by adding a filter
/// for each district string/number in the column or since that may be difficult
/// to figure out programmatically you can set a custom filter. Excel uses both
/// of these methods depending on the data being filtered.
///
/// ```
/// # // This code is available in examples/doc_worksheet_filter_column7.rs
/// #
/// # use rust_xlsxwriter::{FilterCondition, FilterCriteria, Workbook, XlsxError};
/// #
/// # fn main() -> Result<(), XlsxError> {
/// #     let mut workbook = Workbook::new();
/// #
/// #     // Add a worksheet with some sample data to filter.
/// #     let worksheet = workbook.add_worksheet();
/// #     worksheet.write_string(0, 0, "Region")?;
/// #     worksheet.write_string(1, 0, "")?;
/// #     worksheet.write_string(2, 0, "West")?;
/// #     worksheet.write_string(3, 0, "East")?;
/// #     worksheet.write_string(4, 0, "")?;
/// #     worksheet.write_string(5, 0, "")?;
/// #     worksheet.write_string(6, 0, "West")?;
/// #
/// #     worksheet.write_string(0, 1, "Sales")?;
/// #     worksheet.write_number(1, 1, 3000)?;
/// #     worksheet.write_number(2, 1, 8000)?;
/// #     worksheet.write_number(3, 1, 5000)?;
/// #     worksheet.write_number(4, 1, 4000)?;
/// #     worksheet.write_number(5, 1, 7000)?;
/// #     worksheet.write_number(6, 1, 9000)?;
/// #
/// #     // Set the autofilter.
/// #     worksheet.autofilter(0, 0, 6, 1)?;
///
///     // Filter non-blanks by filtering on all the unique non-blank
///     // strings/numbers in the column.
///     let filter_condition = FilterCondition::new()
///         .add_list_filter("East")
///         .add_list_filter("West")
///         .add_list_filter("North")
///         .add_list_filter("South");
///     worksheet.filter_column(0, &filter_condition)?;
///
///     // Or you can add a simpler custom filter to get the same result.
///
///     // Set a custom number filter of `!= " "` to filter non blanks.
///     let filter_condition =
///         FilterCondition::new().add_custom_filter(FilterCriteria::NotEqualTo, " ");
///     worksheet.filter_column(0, &filter_condition)?;
///
/// #     workbook.save("worksheet.xlsx")?;
/// #
/// #     Ok(())
/// # }
/// ```
///
/// Output file:
///
/// <img
/// src="https://rustxlsxwriter.github.io/images/worksheet_filter_column7.png">
///
#[derive(Clone)]
pub struct FilterCondition {
    pub(crate) is_list_filter: bool,
    pub(crate) apply_logical_or: bool,
    pub(crate) should_match_blanks: bool,
    pub(crate) list: Vec<FilterData>,
    pub(crate) custom1: Option<FilterData>,
    pub(crate) custom2: Option<FilterData>,
}

#[allow(clippy::new_without_default)]
impl FilterCondition {
    /// Create a new `FilterCondition` struct to define autofilter rules
    /// associated with an [`autofilter()`](crate::Worksheet::autofilter())
    /// range and the the [`filter_column`](crate::Worksheet::filter_column)
    /// method.
    ///
    /// See the examples above.
    ///
    pub fn new() -> FilterCondition {
        FilterCondition {
            is_list_filter: true,
            apply_logical_or: true,
            should_match_blanks: false,
            list: vec![],
            custom1: None,
            custom2: None,
        }
    }

    /// Add a list filter condition.
    ///
    /// Add a list filter to a column in an autofilter range. This method can be
    /// called multiple times to add multiple "equal to" filter conditions with
    /// a boolean "or". So in the example below the the equivalent Rust
    /// expression for the filter condition would be: `value == "East" || value
    /// == "West" || value == "South"`.
    ///
    /// # Parameters
    ///
    /// - `value`: The value can be a `&str`,`f64` or `i32` type for which the
    ///   [`IntoFilterData`] trait is implemented.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting an autofilter with multiple
    /// list filter conditions.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_filter_column2.rs
    /// #
    /// # use rust_xlsxwriter::{FilterCondition, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with some sample data to filter.
    /// #     let worksheet = workbook.add_worksheet();
    /// #     worksheet.write_string(0, 0, "Region")?;
    /// #     worksheet.write_string(1, 0, "East")?;
    /// #     worksheet.write_string(2, 0, "West")?;
    /// #     worksheet.write_string(3, 0, "East")?;
    /// #     worksheet.write_string(4, 0, "North")?;
    /// #     worksheet.write_string(5, 0, "South")?;
    /// #     worksheet.write_string(6, 0, "West")?;
    /// #
    /// #     worksheet.write_string(0, 1, "Sales")?;
    /// #     worksheet.write_number(1, 1, 3000)?;
    /// #     worksheet.write_number(2, 1, 8000)?;
    /// #     worksheet.write_number(3, 1, 5000)?;
    /// #     worksheet.write_number(4, 1, 4000)?;
    /// #     worksheet.write_number(5, 1, 7000)?;
    /// #     worksheet.write_number(6, 1, 9000)?;
    /// #
    /// #     // Set the autofilter.
    /// #     worksheet.autofilter(0, 0, 6, 1)?;
    /// #
    ///     // Set a filter condition to only show cells matching "East", "West" or
    ///     // "South" in the first column.
    ///     let filter_condition = FilterCondition::new()
    ///         .add_list_filter("East")
    ///         .add_list_filter("West")
    ///         .add_list_filter("South");
    ///     worksheet.filter_column(0, &filter_condition)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_filter_column2.png">
    ///
    pub fn add_list_filter<T>(mut self, value: T) -> FilterCondition
    where
        T: IntoFilterData,
    {
        self.list
            .push(value.new_filter_data(FilterCriteria::EqualTo));
        self.is_list_filter = true;
        self
    }

    /// Add a list filter to filter on Blanks.
    ///
    /// Add a filter condition to a list filter to show Blank cells. For
    /// autofilters Excel treats empty or whitespace only cells as "Blank".
    ///
    /// Filtering non-blanks can be done in two ways. See the second example
    /// below.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting an autofilter with a list
    /// filter for blank cells.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_filter_column3.rs
    /// #
    /// # use rust_xlsxwriter::{FilterCondition, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with some sample data to filter.
    /// #     let worksheet = workbook.add_worksheet();
    /// #     worksheet.write_string(0, 0, "Region")?;
    /// #     worksheet.write_string(1, 0, "")?;
    /// #     worksheet.write_string(2, 0, "West")?;
    /// #     worksheet.write_string(3, 0, "East")?;
    /// #     worksheet.write_string(4, 0, "North")?;
    /// #
    /// #     worksheet.write_string(6, 0, "West")?;
    /// #
    /// #     worksheet.write_string(0, 1, "Sales")?;
    /// #     worksheet.write_number(1, 1, 3000)?;
    /// #     worksheet.write_number(2, 1, 8000)?;
    /// #     worksheet.write_number(3, 1, 5000)?;
    /// #     worksheet.write_number(4, 1, 4000)?;
    /// #     worksheet.write_number(5, 1, 7000)?;
    /// #     worksheet.write_number(6, 1, 9000)?;
    /// #
    /// #     // Set the autofilter.
    /// #     worksheet.autofilter(0, 0, 6, 1)?;
    /// #
    ///     // Set a filter condition to only show cells matching blanks.
    ///     let filter_condition = FilterCondition::new().add_list_blanks_filter();
    ///
    ///     worksheet.filter_column(0, &filter_condition)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_filter_column3.png">
    ///
    ///
    /// The following example demonstrates does the opposite of the previous
    /// example, it sets an autofilter to show all the non-blank values in a
    /// column. This can be done in two ways: by adding a filter for each district
    /// string/number in the column or since that may be difficult to figure out
    /// programmatically you can set a custom filter. Excel uses both of these
    /// methods depending on the data being filtered.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_filter_column7.rs
    /// #
    /// # use rust_xlsxwriter::{FilterCondition, FilterCriteria, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with some sample data to filter.
    /// #     let worksheet = workbook.add_worksheet();
    /// #     worksheet.write_string(0, 0, "Region")?;
    /// #     worksheet.write_string(1, 0, "")?;
    /// #     worksheet.write_string(2, 0, "West")?;
    /// #     worksheet.write_string(3, 0, "East")?;
    /// #     worksheet.write_string(4, 0, "")?;
    /// #     worksheet.write_string(5, 0, "")?;
    /// #     worksheet.write_string(6, 0, "West")?;
    /// #
    /// #     worksheet.write_string(0, 1, "Sales")?;
    /// #     worksheet.write_number(1, 1, 3000)?;
    /// #     worksheet.write_number(2, 1, 8000)?;
    /// #     worksheet.write_number(3, 1, 5000)?;
    /// #     worksheet.write_number(4, 1, 4000)?;
    /// #     worksheet.write_number(5, 1, 7000)?;
    /// #     worksheet.write_number(6, 1, 9000)?;
    /// #
    /// #     // Set the autofilter.
    /// #     worksheet.autofilter(0, 0, 6, 1)?;
    /// #
    ///     // Filter non-blanks by filtering on all the unique non-blank
    ///     // strings/numbers in the column.
    ///     let filter_condition = FilterCondition::new()
    ///         .add_list_filter("East")
    ///         .add_list_filter("West")
    ///         .add_list_filter("North")
    ///         .add_list_filter("South");
    ///     worksheet.filter_column(0, &filter_condition)?;
    ///
    ///     // Or you can add a simpler custom filter to get the same result.
    ///
    ///     // Set a custom number filter of `!= " "` to filter non blanks.
    ///     let filter_condition =
    ///         FilterCondition::new().add_custom_filter(FilterCriteria::NotEqualTo, " ");
    ///     worksheet.filter_column(0, &filter_condition)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_filter_column7.png">
    ///
    pub fn add_list_blanks_filter(mut self) -> FilterCondition {
        self.should_match_blanks = true;
        self.is_list_filter = true;
        self
    }

    /// Add a custom filter condition.
    ///
    /// Add a custom filter to a column in an autofilter range. Excel only
    /// allows two custom conditions so this method can only be called twice.
    ///
    /// When two conditions are specified, like the example below, the logical
    /// operator defaults to "and", like in Excel. However you can use the
    /// [`add_custom_boolean_or`](FilterCondition::add_custom_boolean_or) method
    /// below to get an "or" logical condition.
    ///
    /// # Parameters
    ///
    /// - `value`: The value can be a `&str`,`f64` or `i32` type for which the
    ///   [`IntoFilterData`] trait is implemented.
    /// - `criteria`: The criteria/operator to use in the filter as defined by
    ///   the [`FilterCriteria`] struct.
    ///
    /// # Examples
    ///
    /// The following example demonstrates setting an autofilter for two custom
    /// number filters to create a "between" condition.
    ///
    /// ```
    /// # // This code is available in examples/doc_worksheet_filter_column6.rs
    /// #
    /// # use rust_xlsxwriter::{FilterCondition, FilterCriteria, Workbook, XlsxError};
    /// #
    /// # fn main() -> Result<(), XlsxError> {
    /// #     let mut workbook = Workbook::new();
    /// #
    /// #     // Add a worksheet with some sample data to filter.
    /// #     let worksheet = workbook.add_worksheet();
    /// #     worksheet.write_string(0, 0, "Region")?;
    /// #     worksheet.write_string(1, 0, "East")?;
    /// #     worksheet.write_string(2, 0, "West")?;
    /// #     worksheet.write_string(3, 0, "East")?;
    /// #     worksheet.write_string(4, 0, "North")?;
    /// #     worksheet.write_string(5, 0, "South")?;
    /// #     worksheet.write_string(6, 0, "West")?;
    /// #
    /// #     worksheet.write_string(0, 1, "Sales")?;
    /// #     worksheet.write_number(1, 1, 3000)?;
    /// #     worksheet.write_number(2, 1, 8000)?;
    /// #     worksheet.write_number(3, 1, 5000)?;
    /// #     worksheet.write_number(4, 1, 4000)?;
    /// #     worksheet.write_number(5, 1, 7000)?;
    /// #     worksheet.write_number(6, 1, 9000)?;
    /// #
    /// #     // Set the autofilter.
    /// #     worksheet.autofilter(0, 0, 6, 1)?;
    /// #
    ///     // Set two custom number filters in a "between" configuration.
    ///     let filter_condition = FilterCondition::new()
    ///         .add_custom_filter(FilterCriteria::GreaterThanOrEqualTo, 4000)
    ///         .add_custom_filter(FilterCriteria::LessThanOrEqualTo, 8000);
    ///     worksheet.filter_column(1, &filter_condition)?;
    /// #
    /// #     workbook.save("worksheet.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/worksheet_filter_column6.png">
    ///
    pub fn add_custom_filter<T>(mut self, criteria: FilterCriteria, value: T) -> FilterCondition
    where
        T: IntoFilterData,
    {
        if self.custom1.is_none() {
            self.custom1 = Some(value.new_filter_data(criteria));
        } else if self.custom2.is_none() {
            self.custom2 = Some(value.new_filter_data(criteria));
            self.apply_logical_or = false;
        } else {
            eprintln!("Excel only allows 2 custom filter conditions.");
        }

        self.is_list_filter = false;
        self
    }

    /// Add an "or" logical condition for two custom filters.
    ///
    /// When two conditions are specified, like the example above, the logical
    /// operator defaults to "and", as in Excel. However you can use the
    /// [`add_custom_boolean_or`](FilterCondition::add_custom_boolean_or) method
    /// to get an "or" logical condition.
    ///
    pub fn add_custom_boolean_or(mut self) -> FilterCondition {
        self.apply_logical_or = true;
        self.is_list_filter = false;
        self
    }
}

/// The `FilterCriteria` enum defines logical filter criteria used in an
/// autofilter.
///
/// These filter criteria are used with the [`FilterCondition`]
/// [`add_custom_filter()`](FilterCondition::add_custom_filter) method.
///
/// Currently only Excel's string and number filter operations are supported.
/// The numeric style criteria such as `>=` can also be applied to strings (like
/// in Rust) but the string operations like `BeginsWith` are only applied to
/// strings in Excel.
///
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum FilterCriteria {
    /// Show numbers or strings that are equal to the filter value.
    EqualTo,

    /// Show numbers or strings that are not equal to the filter value.
    NotEqualTo,

    /// Show numbers or strings that are greater than the filter value.
    GreaterThan,

    /// Show numbers or strings that are greater than or equal to the filter value.
    GreaterThanOrEqualTo,

    /// Show numbers or strings that are less than the filter value.
    LessThan,

    /// Show numbers or strings that are less than or equal to the filter value.
    LessThanOrEqualTo,

    /// Show strings that begin with the filter string value.
    BeginsWith,

    /// Show strings that do not begin with the filter string value.
    DoesNotBeginWith,

    /// Show strings that end with the filter string value.
    EndsWith,

    /// Show strings that do not end with the filter string value.
    DoesNotEndWith,

    /// Show strings that contain with the filter string value.
    Contains,

    /// Show strings that do not contain with the filter string value.
    DoesNotContain,
}

#[allow(clippy::match_same_arms)]
impl FilterCriteria {
    pub(crate) fn operator(self) -> String {
        match self {
            FilterCriteria::EqualTo => String::new(),
            FilterCriteria::LessThan => "lessThan".to_string(),
            FilterCriteria::NotEqualTo => "notEqual".to_string(),
            FilterCriteria::GreaterThan => "greaterThan".to_string(),
            FilterCriteria::LessThanOrEqualTo => "lessThanOrEqual".to_string(),
            FilterCriteria::GreaterThanOrEqualTo => "greaterThanOrEqual".to_string(),
            FilterCriteria::EndsWith => String::new(),
            FilterCriteria::Contains => String::new(),
            FilterCriteria::BeginsWith => String::new(),
            FilterCriteria::DoesNotEndWith => "notEqual".to_string(),
            FilterCriteria::DoesNotContain => "notEqual".to_string(),
            FilterCriteria::DoesNotBeginWith => "notEqual".to_string(),
        }
    }
}

/// The FilterData struct represents data types used in Excel's filters.
///
/// The FilterData struct is a simple data type to allow a generic mapping
/// between Rust's string and number types and similar types used in Excel's
/// filters.
#[doc(hidden)]
#[derive(Clone)]
pub struct FilterData {
    pub(crate) data_type: FilterDataType,
    pub(crate) string: String,
    pub(crate) number: f64,
    pub(crate) criteria: FilterCriteria,
}

impl FilterData {
    fn new_string_and_criteria(value: &str, criteria: FilterCriteria) -> FilterData {
        FilterData {
            data_type: FilterDataType::String,
            string: value.to_string(),
            number: 0.0,
            criteria,
        }
    }

    fn new_number_and_criteria(value: f64, criteria: FilterCriteria) -> FilterData {
        // Store number but also convert it to a string since Excel makes string
        // comparisons to "numbers stored as strings".
        FilterData {
            data_type: FilterDataType::Number,
            string: value.to_string(),
            number: value,
            criteria,
        }
    }

    // Excel stores some of the string operators as simple regex patterns.
    pub(crate) fn value(&self) -> String {
        match self.criteria {
            FilterCriteria::EndsWith | FilterCriteria::DoesNotEndWith => {
                format!("*{}", self.string)
            }
            FilterCriteria::Contains | FilterCriteria::DoesNotContain => {
                format!("*{}*", self.string)
            }
            FilterCriteria::BeginsWith | FilterCriteria::DoesNotBeginWith => {
                format!("{}*", self.string)
            }
            // For everything else, including numbers, we just use the string value.
            _ => self.string.clone(),
        }
    }
}

/// Trait to map different Rust types into Excel data types used in filters.
///
/// Currently only string and number like types are supported.
pub trait IntoFilterData {
    /// Types/objects supporting this trait must be able to convert to a
    /// `FilterData` struct.
    fn new_filter_data(&self, criteria: FilterCriteria) -> FilterData;
}

impl IntoFilterData for f64 {
    fn new_filter_data(&self, criteria: FilterCriteria) -> FilterData {
        FilterData::new_number_and_criteria(*self, criteria)
    }
}

impl IntoFilterData for i32 {
    fn new_filter_data(&self, criteria: FilterCriteria) -> FilterData {
        FilterData::new_number_and_criteria(f64::from(*self), criteria)
    }
}

impl IntoFilterData for &str {
    fn new_filter_data(&self, criteria: FilterCriteria) -> FilterData {
        FilterData::new_string_and_criteria(self, criteria)
    }
}

#[derive(Clone, PartialEq, Eq)]
pub(crate) enum FilterDataType {
    String,
    Number,
}
