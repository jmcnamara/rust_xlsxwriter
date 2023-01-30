// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting an autofilter to show all the
//! non-blank values in a column. This can be done in 2 ways: by adding a filter
//! for each district string/number in the column or since that may be difficult
//! to figure out programmatically you can set a custom filter. Excel uses both
//! of these methods depending on the data being filtered.

use rust_xlsxwriter::{FilterCondition, FilterCriteria, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    worksheet.write_string_only(0, 0, "Region")?;
    worksheet.write_string_only(1, 0, "")?;
    worksheet.write_string_only(2, 0, "West")?;
    worksheet.write_string_only(3, 0, "East")?;
    worksheet.write_string_only(4, 0, "")?;
    worksheet.write_string_only(5, 0, "")?;
    worksheet.write_string_only(6, 0, "West")?;

    worksheet.write_string_only(0, 1, "Sales")?;
    worksheet.write_number_only(1, 1, 3000)?;
    worksheet.write_number_only(2, 1, 8000)?;
    worksheet.write_number_only(3, 1, 5000)?;
    worksheet.write_number_only(4, 1, 4000)?;
    worksheet.write_number_only(5, 1, 7000)?;
    worksheet.write_number_only(6, 1, 9000)?;

    // Set the autofilter.
    worksheet.autofilter(0, 0, 6, 1)?;

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

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
