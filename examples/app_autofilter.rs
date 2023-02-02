// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of how to create autofilters with the rust_xlsxwriter library.
//!
//! An autofilter is a way of adding drop down lists to the headers of a 2D
//! range of worksheet data. This allows users to filter the data based on
//! simple criteria so that some data is shown and some is hidden.
//!

use rust_xlsxwriter::{FilterCondition, FilterCriteria, Format, Workbook, Worksheet, XlsxError};

fn main() -> Result<(), XlsxError> {
    // Create a new Excel file object.
    let mut workbook = Workbook::new();

    // -----------------------------------------------------------------------
    // 1. Add an autofilter to a data range.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, false)?;

    // Set the autofilter area, including the header/filter row.
    worksheet.autofilter(0, 0, 50, 3)?;

    // -----------------------------------------------------------------------
    // 2. Add an autofilter with a list filter condition.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, false)?;

    // Set the autofilter area.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Set a filter condition to only show cells matching "East" in the first
    // column.
    let filter_condition = FilterCondition::new().add_list_filter("East");
    worksheet.filter_column(0, &filter_condition)?;

    // -----------------------------------------------------------------------
    // 3. Add an autofilter with a list filter condition on multiple items.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, false)?;

    // Set the autofilter area.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Set a filter condition to only show cells matching "East", "West" or
    // "South" in the first column.
    let filter_condition = FilterCondition::new()
        .add_list_filter("East")
        .add_list_filter("West")
        .add_list_filter("South");
    worksheet.filter_column(0, &filter_condition)?;

    // -----------------------------------------------------------------------
    // 4. Add an autofilter with a list filter condition to match blank cells.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, true)?;

    // Set the autofilter area.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Set a filter condition to only show cells matching blanks.
    let filter_condition = FilterCondition::new().add_list_blanks_filter();
    worksheet.filter_column(0, &filter_condition)?;

    // -----------------------------------------------------------------------
    // 5. Add an autofilter with list filters in multiple columns.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, false)?;

    // Set the autofilter area.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Set a filter condition for 2 separate columns.
    let filter_condition1 = FilterCondition::new().add_list_filter("East");
    worksheet.filter_column(0, &filter_condition1)?;

    let filter_condition2 = FilterCondition::new().add_list_filter("July");
    worksheet.filter_column(3, &filter_condition2)?;

    // -----------------------------------------------------------------------
    // 6. Add an autofilter with custom filter condition.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, false)?;

    // Set the autofilter area for numbers greater than 8000.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Set a custom number filter.
    let filter_condition =
        FilterCondition::new().add_custom_filter(FilterCriteria::GreaterThan, 8000);
    worksheet.filter_column(2, &filter_condition)?;

    // -----------------------------------------------------------------------
    // 7. Add an autofilter with 2 custom filters to create a "between" condition.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, false)?;

    // Set the autofilter area.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Set two custom number filters in a "between" configuration.
    let filter_condition = FilterCondition::new()
        .add_custom_filter(FilterCriteria::GreaterThanOrEqualTo, 4000)
        .add_custom_filter(FilterCriteria::LessThanOrEqualTo, 6000);
    worksheet.filter_column(2, &filter_condition)?;

    // -----------------------------------------------------------------------
    // 8. Add an autofilter for non blanks.
    // -----------------------------------------------------------------------

    // Add a worksheet  with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    populate_autofilter_data(worksheet, true)?;

    // Set the autofilter area.
    worksheet.autofilter(0, 0, 50, 3)?;

    // Filter non-blanks by filtering on all the unique non-blank
    // strings/numbers in the column.
    let filter_condition = FilterCondition::new()
        .add_list_filter("East")
        .add_list_filter("West")
        .add_list_filter("North")
        .add_list_filter("South");
    worksheet.filter_column(0, &filter_condition)?;

    // Or you can add a simpler custom filter to get the same result.

    // Set a custom number filter of `!= " "` to filter non blanks.
    let filter_condition =
        FilterCondition::new().add_custom_filter(FilterCriteria::NotEqualTo, " ");
    worksheet.filter_column(0, &filter_condition)?;

    // Save the file to disk.
    workbook.save("autofilter.xlsx")?;

    Ok(())
}

// Generate worksheet data to filter on.
pub fn populate_autofilter_data(
    worksheet: &mut Worksheet,
    add_blanks: bool,
) -> Result<(), XlsxError> {
    // The sample data to add to the worksheet.
    let mut data = vec![
        ("East", "Apple", 9000, "July"),
        ("East", "Apple", 5000, "April"),
        ("South", "Orange", 9000, "September"),
        ("North", "Apple", 2000, "November"),
        ("West", "Apple", 9000, "November"),
        ("South", "Pear", 7000, "October"),
        ("North", "Pear", 9000, "August"),
        ("West", "Orange", 1000, "December"),
        ("West", "Grape", 1000, "November"),
        ("South", "Pear", 10000, "April"),
        ("West", "Grape", 6000, "January"),
        ("South", "Orange", 3000, "May"),
        ("North", "Apple", 3000, "December"),
        ("South", "Apple", 7000, "February"),
        ("West", "Grape", 1000, "December"),
        ("East", "Grape", 8000, "February"),
        ("South", "Grape", 10000, "June"),
        ("West", "Pear", 7000, "December"),
        ("South", "Apple", 2000, "October"),
        ("East", "Grape", 7000, "December"),
        ("North", "Grape", 6000, "July"),
        ("East", "Pear", 8000, "February"),
        ("North", "Apple", 7000, "August"),
        ("North", "Orange", 7000, "July"),
        ("North", "Apple", 6000, "June"),
        ("South", "Grape", 8000, "September"),
        ("West", "Apple", 3000, "October"),
        ("South", "Orange", 10000, "November"),
        ("West", "Grape", 4000, "December"),
        ("North", "Orange", 5000, "August"),
        ("East", "Orange", 1000, "November"),
        ("East", "Orange", 4000, "October"),
        ("North", "Grape", 5000, "August"),
        ("East", "Apple", 1000, "July"),
        ("South", "Apple", 10000, "March"),
        ("East", "Grape", 7000, "October"),
        ("West", "Grape", 1000, "September"),
        ("East", "Grape", 10000, "October"),
        ("South", "Orange", 8000, "March"),
        ("North", "Apple", 4000, "July"),
        ("South", "Orange", 5000, "July"),
        ("West", "Apple", 4000, "June"),
        ("East", "Apple", 5000, "April"),
        ("North", "Pear", 3000, "August"),
        ("East", "Grape", 9000, "November"),
        ("North", "Orange", 8000, "October"),
        ("East", "Apple", 10000, "June"),
        ("South", "Pear", 1000, "December"),
        ("North", "Grape", 10000, "July"),
        ("East", "Grape", 6000, "February"),
    ];

    // Introduce blanks cells for some of the examples.
    if add_blanks {
        data[5].0 = "";
        data[18].0 = "";
        data[30].0 = "";
        data[40].0 = "";
    }

    // Widen the columns for clarity.
    worksheet.set_column_width(0, 12)?;
    worksheet.set_column_width(1, 12)?;
    worksheet.set_column_width(2, 12)?;
    worksheet.set_column_width(3, 12)?;

    // Write the header titles.
    let header_format = Format::new().set_bold();
    worksheet.write_string_with_format(0, 0, "Region", &header_format)?;
    worksheet.write_string_with_format(0, 1, "Item", &header_format)?;
    worksheet.write_string_with_format(0, 2, "Volume", &header_format)?;
    worksheet.write_string_with_format(0, 3, "Month", &header_format)?;

    // Write the other worksheet data.
    for (row, data) in data.iter().enumerate() {
        let row = 1 + row as u32;
        worksheet.write_string(row, 0, data.0)?;
        worksheet.write_string(row, 1, data.1)?;
        worksheet.write_number(row, 2, data.2)?;
        worksheet.write_string(row, 3, data.3)?;
    }

    Ok(())
}
