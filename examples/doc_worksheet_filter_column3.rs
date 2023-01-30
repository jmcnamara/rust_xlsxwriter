// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! The following example demonstrates setting an autofilter with a list filter
//! for blank cells.

use rust_xlsxwriter::{FilterCondition, Workbook, XlsxError};

fn main() -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();

    // Add a worksheet with some sample data to filter.
    let worksheet = workbook.add_worksheet();
    worksheet.write_string_only(0, 0, "Region")?;
    worksheet.write_string_only(1, 0, "")?;
    worksheet.write_string_only(2, 0, "West")?;
    worksheet.write_string_only(3, 0, "East")?;
    worksheet.write_string_only(4, 0, "North")?;

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

    // Set a filter condition to only show cells matching blanks.
    let filter_condition = FilterCondition::new().add_list_blanks_filter();

    worksheet.filter_column(0, &filter_condition)?;

    workbook.save("worksheet.xlsx")?;

    Ok(())
}
